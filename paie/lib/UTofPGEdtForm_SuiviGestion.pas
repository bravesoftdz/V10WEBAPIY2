{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 08/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE :  ()
Mots clefs ... : TOF;EDITIONS FORMATION
Contient les sources TOF pour les éditions de la formation
*****************************************************************
PT1 | 24/04/2003 | V_42  | JL | Développement pour CWAS
PT2 | 05/11/2003 | V_50  | JL | Correction erreur type variant incorretc pour suivi financier
PT3 | 19/11/2003 | V_50  | JL | Correction affichage rupture par type investissement
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT4 | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
    |            |       |    | pour les critères code salarié uniquement
PT5 | 03/04/2008 | V_80  | FL | Adaptation pour partage formation + plus de lanceEtatTob 
PT6 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue + Gestion elipsis salarié et responsable
}

unit UTofPGEdtForm_SuiviGestion;

interface
Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,Controls,Classes, UTOB,
{$IFDEF EAGLCLIENT}
     eQRS1,UtilEAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtREtat,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,LookUp,
     PGOutilsFormation,P5Def,PGEditOutils,PGEditOutils2,HQry,HTB97,ParamDat,
     PGEdtEtat,pgOutils ;

Type
  TOF_PGEDITFORGESTIND = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate ; override ;
    private
    OrderBySal,OrderByFor:String;
    procedure SalarieElipsisClick(Sender : TObject);
    procedure RuptureSal(Sender : TObject);
    procedure RuptureFor(Sender : TObject);
    procedure RespElipsisClick(Sender : TObject);
  end ;

  TOF_PGEDITFORGESTPREV  = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate ; override ;
    private
    OrderBySal:String;
    procedure SalarieElipsisClick(Sender : TObject);
    procedure RuptureSal(Sender : TObject);
  end ;

  TOF_PGFORSUIVIGESTCOLL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate ; override ;
    private
    OrderBy,GroupBy:String;
    procedure CreerRupture(Champ,Tablette : String;Num : Integer);
  end;

  TOF_PGEDITFORSUIVIINV = Class (TOF)
    procedure OnArgument (S : String ) 	; override ;
    procedure OnUpdate 					; override ;
    procedure OnClose  					; override ; //PT5
    private
    TobEtat : TOB; //PT5
    procedure DateElipsisClick (Sender : TObject);
    procedure EtatAvecTob; // (Sender : TObject); //PT5
    Function ChampRuptureForm (ValeurRupt : String) : String;
	procedure SalarieElipsisClick(Sender : TObject); //PT5
	procedure RespElipsisClick(Sender : TObject); //PT5
//	procedure StageElipsisClick(Sender : TObject); //PT5
  end;

  TOF_PGFORCOMPCOLL = Class (TOF)
    procedure OnArgument(S:String); override;
    private
    procedure SalarieElipsisClick(Sender : TObject);
    procedure BEtatClick(Sender : TObject);
    Function ChampRuptureForm(TypeRupt,ValeurRupt : String) : String;
    procedure ChangeMillesime (Sender : TObject);
    end;
  TOF_PGEDITFORCOMPDETAIL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate ; override ;
    end;

implementation
{TOF_PGEDITFORGESTIND}
procedure TOF_PGEDITFORGESTIND.OnArgument (S : String ) ;
var Num : Integer;
    Min,Max,Arg : String;
    Combo : THValComboBox;
    Check : TCheckBox;
    Resp,Edit : THEdit;
    DD,DF : TDateTime;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        // Debut PT1
        If Arg ='CWAS' then
        begin
                SetControlVisible('PFO_RESPONSFOR',False);
                SetControlVisible('TRESP',False);
                SetControlVisible('LIBRESP',False);
                SetControlVisible('CRUPTRESP',False);
                SetControlText('PFO_RESPONSFOR',V_PGI.UserSalarie);
                Edit := THEdit(GetControl('PFO_SALARIE'));
                If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
                Edit := THEdit(GetControl('PFO_SALARIE_'));
                If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        end
        else
        begin
                RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
                SetcontrolText('PFO_SALARIE',Min);
                SetControlText('PFO_SALARIE_',Max);
        end;
        // Fin PT1
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        For Num := 1 To 8 do
        begin
                Combo := THValComboBox(GetControl('PST_FORMATION'+IntToStr(Num)));
                If Combo.Visible=True Then SetControlVisible('CRUPT'+IntToStr(Num),True);
                Check := TCheckBox(GetControl('CRUPT'+IntToStr(Num)));
                If Check <> NIl Then Check.OnClick := RuptureFor;
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
                SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
        end;
        VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
        Combo := THValComboBox(GetControl('PFO_CODESTAT'));
        If Combo.Visible=true then SetControlVisible('CRUPTSAL5',True);

        RendMillesimeRealise(DD,DF);
        SetControlText('PFO_DATEDEBUT',DateToStr(DD));
        SetControlText('PFO_DATEFIN',dateToStr(DF));
        SetcontrolCaption('LIBRESP','');
        SetcontrolCaption('LIBSTAGE','');
        OrderBySal := '';
        OrderByFor := '';
        For Num := 1 to 5 do
        begin
                Check := TCheckBox(GetControl('CRUPTSAL'+IntToStr(Num)));
                If Check <> Nil then Check.OnClick := RuptureSal;
        end;
        Check := TCheckBox(GetControl('CRUPTETAB'));
        If Check <> Nil then Check.OnClick := RuptureSal;
        Check := TCheckBox(GetControl('CRUPTRESP'));
        If Check <> Nil then Check.OnClick := RuptureSal;
        If PGBundleHierarchie Then //PT6
        Begin
	        Resp := THEdit(GetControl('PFO_RESPONSFOR'));
	        If Resp <> Nil Then Resp.OnElipsisClick := RespElipsisClick;
	    End;
end;

procedure TOF_PGEDITFORGESTIND.OnUpdate ;
var Where,SQL,OrderBy : String;
    Pages : TPageControl;
begin
  Inherited ;
        Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        If Where <> '' Then Where := ' '+Where+' AND PFO_EFFECTUE="X"'
        Else Where := ' WHERE PFO_EFFECTUE="X"';
        If OrderBySal <> '' Then OrderBy := ' ORDER BY '+OrderBySal+','
        Else OrderBy := ' ORDER BY ';
        // Rupture 2 : Matricule ou nom salarié

        If GetControlText('CALPHA') = 'X' Then
        begin
                SetControlText('XX_RUPTURE2','PFO_SALARIE');
                OrderBy := orderBy+'PFO_NOMSALARIE,PFO_SALARIE';
        end
        Else
        begin
                SetControlText('XX_RUPTURE2','PFO_SALARIE');
                OrderBy := OrderBy+'PFO_SALARIE';
        end;
        If OrderByFor <> '' Then OrderBy := OrderBy+','+OrderByFor;
        SQL := 'SELECT PFO_NBREHEURE,PFO_RESPONSFOR,PFO_SALARIE,PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_MILLESIME,PFO_ORDRE,'+
        'PFO_ETABLISSEMENT,PFO_PRENOM,PFO_NOMSALARIE,PFO_TRAVAILN1,PFO_TRAVAILN2,PFO_TRAVAILN3,PFO_TRAVAILN4,PFO_CODESTAT,'+
        'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4'+
        ',PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PSA_DATENAISSANCE,PSA_DATEENTREE FROM FORMATIONS'+
        ' LEFT JOIN SESSIONSTAGE ON PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME AND PFO_CODESTAGE=PSS_CODESTAGE'+
        ' LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE'+
        ' LEFT JOIN SALARIES SALARIES ON PFO_SALARIE=PSA_SALARIE'+Where+OrderBy;
        TFQRS1(Ecran).WhereSQL := SQL;
end;

procedure TOF_PGEDITFORGESTIND.SalarieElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        StOrder := 'PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT4
        LookupList(THEdit(Sender),'Liste des stages','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;

procedure TOF_PGEDITFORGESTIND.RuptureSal(Sender : TObject);
var i,a : Integer;
begin
        If TCheckBox(Sender).Name = 'CRUPTETAB' Then
        begin
                If TCheckBox(SENDER).Checked=True Then
                begin
                        SetControlChecked('CRUPTRESP',False);
                        For i := 1 To 5 do SetControlChecked('CRUPTSAL'+IntToStr(i),False);
                        OrderBySal := 'PFO_ETABLISSEMENT';
                end
                Else OrderBySal := '';
                SetControlText('XX_RUPTURE1',OrderBySal);
                Exit;
        end;
        For i := 1 to 5 do
        begin
                If TCheckBox(Sender).Name = 'CRUPTSAL'+IntToStr(i) Then
                begin
                        If TCheckBox(SENDER).Checked=True Then
                        begin
                                SetControlChecked('CRUPTRESP',False);
                                SetControlChecked('CRUPTETAB',False);
                                For a := 1 To 5 do
                                begin
                                        If a <> i Then SetControlChecked('CRUPTSAL'+IntToStr(a),False);
                                end;
                                If i<5 Then OrderBySal := 'PFO_TRAVAILN'+IntToStr(i)
                                Else OrderBySal := 'PFO_CODESTAT';
                        end
                        Else OrderBySal := '';
                end;
        end;
        If TCheckBox(Sender).Name = 'CRUPTRESP' Then
        begin
                If TCheckBox(SENDER).Checked=True Then
                begin
                        SetControlChecked('CRUPTETAB',False);
                        For i := 1 To 5 do SetControlChecked('CRUPTSAL'+IntToStr(i),False);
                        OrderBySal := 'PFO_RESPONSFOR';
                end
                Else OrderBySal := '';
                SetControlText('XX_RUPTURE1',OrderBySal);
                Exit;
        end;
        SetControlText('XX_RUPTURE1',OrderBySal);
end;

procedure TOF_PGEDITFORGESTIND.RuptureFor(Sender : TObject);
var a,i : Integer;
begin
        For i := 1 to 5 do
        begin
                If TCheckBox(Sender).Name = 'CRUPT'+IntToStr(i) Then
                begin
                        If TCheckBox(SENDER).Checked=True Then
                        begin
                                For a := 1 To 5 do
                                begin
                                        If a <> i Then SetControlChecked('CRUPT'+IntToStr(a),False);
                                end;
                                OrderByFor := 'PST_FORMATION'+IntToStr(i);
                        end
                        Else OrderByFor := '';
                end;
        end;
        SetControlText('XX_RUPTURE3',OrderByFor);
end;

procedure TOF_PGEDITFORGESTIND.RespElipsisClick(Sender : TObject);
//var StWhere : String;
begin
	//PT5
        //StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES)';
        //LookupList(THEdit(Sender),'Liste des responsables de formation','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,'', True,-1);
        ElipsisResponsableMultidos(Sender);
end;

{TOF_PGEDITFORGESTPREV}

procedure TOF_PGEDITFORGESTPREV.OnArgument (S : String ) ;
var Num : Integer;
    Millesime : THValComboBox;
    Check : TCheckBox;
    Arg : String;
    Edit : THEdit;
    DD,DF : TDateTime;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        //Debut PT1
        If Arg ='CWAS' then
        begin
                SetControlVisible('PSE_RESPONSFOR',False);
                SetControlVisible('TRESPONSABLE',False);
                SetControlVisible('CRUPTRESP',False);
                SetControlVisible('LIBRESP',False);
                SetControlText('PSE_RESPONSFOR',V_PGI.UserSalarie);
                Edit := THEdit(GetControl('PSA_SALARIE'));
                If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
                Edit := THEdit(GetControl('PSA_SALARIE_'));
                If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        end;
        // Fin PT1
        OrderBySal := '';
        SetControlText('XX_RUPTURE2','PSA_SALARIE');
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
                SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
        Millesime := THValComboBox(GetControl('MILLESIME'));
        Millesime.Value := RendMillesimeRealise(DD,DF);
        SetControlText('DATEDEBUT',DateToStr(DD));
        SetControlText('DATEFIN',DateToStr(DF));
        SetcontrolCaption('LIBRESP','');
        SetcontrolCaption('LIBSTAGE','');
        For Num := 1 to 5 do
        begin
                Check := TCheckBox(GetControl('CRUPTSAL'+IntToStr(Num)));
                If Check <> Nil then Check.OnClick := RuptureSal;
        end;
        Check := TCheckBox(GetControl('CRUPTETAB'));
        If Check <> Nil then Check.OnClick := RuptureSal;
        Check := TCheckBox(GetControl('CRUPTRESP'));
        If Check <> Nil then Check.OnClick := RuptureSal;
        SetControlVisible('CRUPTSTAGE',False);
        SetControlVisible('CRUPTRESP',False);
        SetControlVisible('CRUPTETAB',False);
        SetControlVisible('CRUPTSAL1',False);
        SetControlVisible('CRUPTSAL2',False);
        SetControlVisible('CRUPTSAL3',False);
        SetControlVisible('CRUPTSAL4',False);
        SetControlVisible('CRUPTSAL5',False);
end;

procedure TOF_PGEDITFORGESTPREV.OnUpdate ;
var Pages : TpageControl;
    OrderBy,Where,SQL,WhereFormation : String;
    i : Integer;
    Combo,Millesime : THValComboBox;
    DateDebut,DateFin : TDateTime;
    Q : TQuery;
begin
  Inherited ;
        If GetCheckBoxState('CSAUT4')=CbChecked Then SetControlChecked('CSAUT3',True);
        If GetCheckBoxState('CSAUT3')=CbChecked Then SetControlChecked('CSAUT2',True);
        If GetCheckBoxState('CSAUT2')=CbChecked Then SetControlChecked('CSAUT1',True);
        Millesime := THValComboBox(GetControl('MILLESIME'));
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime.Value+'"',true);
        If Not Q.eof Then
        begin
                DateDebut := Q.FindField('PFE_DATEDEBUT').AsDateTime;
                DateFin := Q.FindField('PFE_DATEFIN').AsDateTime;
        end
        Else begin DateDebut := IDate1900;Datefin := IDate1900;end;
        Ferme(Q);
        OrderBy := '';
        If GetControlText('CALPHA') = 'X' Then
        begin
                If OrderBySal <> '' Then OrderBy := 'ORDER BY '+OrderBySal+',PSA_LIBELLE,PSA_SALARIE'
                Else OrderBy := 'ORDER BY PSA_LIBELLE,PSA_SALARIE';
        end
        Else
        begin
                If OrderBySal <> '' Then OrderBy := 'ORDER BY '+OrderBySal+',PSA_SALARIE'
                Else OrderBy := 'ORDER BY PSA_SALARIE';
        end;
        SetControlText('DATEDEBUT',DateToStr(DateDebut));
        SetControlText('DATEFIN',DateToStr(DateFin));
        Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        If Where <> '' Then Where := ' '+Where+' AND '
        Else Where := ' WHERE ';
        WhereFormation := '';
        For i := 1 to 8 do
        begin
                Combo := THValComboBox(GetControl('FORMATION'+IntToStr(i)));
                If Combo.Value <> '' Then WhereFormation := WhereFormation+' AND PST_FORMATION'+IntToStr(i)+'="'+Combo.Value+'"';
        end;
        If GetControlText('CODESTAGE') <> '' Then WhereFormation := WhereFormation+' AND PST_CODESTAGE="'+GetControlText('CODESTAGE')+'"';
        SQL := 'SELECT PSE_RESPONSFOR,PSE_SALARIE,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_CODESTAT,PSA_ETABLISSEMENT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4 '+
        'FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE'+Where+
        '(EXISTS(SELECT PFO_SALARIE FROM FORMATIONS LEFT JOIN STAGE ON PFO_MILLESIME=PST_MILLESIME AND PFO_CODESTAGE=PST_CODESTAGE '+
        'WHERE PFO_EFFECTUE="X" AND PSA_SALARIE=PFO_SALARIE AND PFO_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'"'+WhereFormation+')'+
        ' OR EXISTS (SELECT PFI_SALARIE FROM INSCFORMATION LEFT JOIN STAGE ON PFI_MILLESIME=PST_MILLESIME AND PFI_CODESTAGE=PST_CODESTAGE '+
        'WHERE PFI_SALARIE=PSA_SALARIE AND PFI_MILLESIME="'+Millesime.Value+'"'+WhereFormation+')) '+OrderBy;
        TFQRS1(Ecran).WhereSQL := SQL;
end;

procedure TOF_PGEDITFORGESTPREV.SalarieElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        StOrder := 'PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT4
        LookupList(THEdit(Sender),'Liste des stages','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;

procedure TOF_PGEDITFORGESTPREV.RuptureSal(Sender : TObject);
var i,a : Integer;
begin
        If TCheckBox(Sender).Name = 'CRUPTETAB' Then
        begin
                If TCheckBox(SENDER).Checked=True Then
                begin
                        SetControlChecked('CRUPTRESP',False);
                        For i := 1 To 5 do SetControlChecked('CRUPTSAL'+IntToStr(i),False);
                        OrderBySal := 'PSA_ETABLISSEMENT';
                end
                Else OrderBySal := '';
                SetControlText('XX_RUPTURE1',OrderBySal);
                Exit;
        end;
        For i := 1 to 5 do
        begin
                If TCheckBox(Sender).Name = 'CRUPTSAL'+IntToStr(i) Then
                begin
                        If TCheckBox(SENDER).Checked=True Then
                        begin
                                SetControlChecked('CRUPTRESP',False);
                                SetControlChecked('CRUPTETAB',False);
                                For a := 1 To 5 do
                                begin
                                        If a <> i Then SetControlChecked('CRUPTSAL'+IntToStr(a),False);
                                end;
                                If i<5 Then OrderBySal := 'PSA_TRAVAILN'+IntToStr(i)
                                Else OrderBySal := 'PSA_CODESTAT';
                        end
                        Else OrderBySal := '';
                end;
        end;
        If TCheckBox(Sender).Name = 'CRUPTRESP' Then
        begin
                If TCheckBox(SENDER).Checked=True Then
                begin
                        SetControlChecked('CRUPTETAB',False);
                        For i := 1 To 5 do SetControlChecked('CRUPTSAL'+IntToStr(i),False);
                        OrderBySal := 'PSE_RESPONSFOR';
                end
                Else OrderBySal := '';
                SetControlText('XX_RUPTURE1',OrderBySal);
                Exit;
        end;
        SetControlText('XX_RUPTURE1',OrderBySal);
end;

{TOF_PGFORSUIVIGESTCOLL}
procedure TOF_PGFORSUIVIGESTCOLL.OnArgument (S : String ) ;
var Combo : THValComboBox;
    Num,i : Integer;
    DD, DF : TDateTime;
begin
  Inherited ;
        SetControlCaption('LIBSTAGE','');
        SetControlCaption('LIBRESP','');
        RendMillesimeRealise(DD,DF);
        SetControlText('PFO_DATEDEBUT',DateToStr(DD));
        SetControlText('PFO_DATEFIN',DateToStr(DF));
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
                SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
        end;
        VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('THVRUPTURE'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                        If i <> 1 Then
                        begin
                                Combo.Items.Add('<<Aucun>>');
                                Combo.Values.Add('');
                        end;
                        Combo.Items.Add('Responsable formation');
                        Combo.Values.Add('RF');
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('FO');
                        Combo.Items.Add('Etablissement');
                        Combo.Values.Add('ET');
                        Combo.Items.Add('Libellé emploi');
                        Combo.Values.Add('LE');
                        Combo.Items.Add('Type (Interne/Externe)');
                        Combo.Values.Add('TY');
                        For Num  :=  1 to VH_Paie.NBFormationLibre do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('F1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('F2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('F3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('F4');end;
                                If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('F5');end;
                                If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('F6');end;
                                If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('F7');end;
                                If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('F8');end;
                        end;
                        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('S1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('S2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('S3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('S4');end;
                        end;
                        If VH_Paie.PGLibCodeStat <> '' Then begin Combo.Items.Add(VH_Paie.PGLibCodeStat); Combo.Values.Add('ST');end;
                end;
        end;
        SetControlProperty('THVRUPTURE1','Value','FO');
        SetControlProperty('THVRUPTURE2','Value','');
        SetControlProperty('THVRUPTURE3','Value','');
        SetControlProperty('THVRUPTURE4','Value','');
end;

procedure TOF_PGFORSUIVIGESTCOLL.OnUpdate;
var Combo : THValComboBox;
    SQL,Where : String;
    Pages : TPageControl;
    i : Integer;
begin
  Inherited ;
        Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        OrderBy := '';
        GroupBy := '';
        For i := 1 to 4 do
        begin
                SetControltext('XX_RUPTURE'+IntToStr(i),'');
                SetControltext('TABLETTE'+IntToStr(i),'');
                SetControltext('CHAMP'+IntToStr(i),'');
        end;
        For i := 1 to 4 do
        begin
                Combo := THValComboBox(GetControl('THVRUPTURE'+IntToStr(i)));
                If Combo.Value <> '' Then
                begin
                        If Combo.Value = 'F1' Then CreerRupture('PST_FORMATION1',VH_Paie.FormationLibre1,i);
                        If Combo.Value = 'F2' Then CreerRupture('PST_FORMATION2',VH_Paie.FormationLibre2,i);
                        If Combo.Value = 'F3' Then CreerRupture('PST_FORMATION3',VH_Paie.FormationLibre3,i);
                        If Combo.Value = 'F4' Then CreerRupture('PST_FORMATION4',VH_Paie.FormationLibre4,i);
                        If Combo.Value = 'F5' Then CreerRupture('PST_FORMATION5',VH_Paie.FormationLibre5,i);
                        If Combo.Value = 'F6' Then CreerRupture('PST_FORMATION6',VH_Paie.FormationLibre6,i);
                        If Combo.Value = 'F7' Then CreerRupture('PST_FORMATION7',VH_Paie.FormationLibre7,i);
                        If Combo.Value = 'F8' Then CreerRupture('PST_FORMATION8',VH_Paie.FormationLibre8,i);
                        If Combo.Value = 'ET' Then CreerRupture('PFO_ETABLISSEMENT','Etablissement',i);
                        If Combo.Value = 'LE' Then CreerRupture('PFO_LIBEMPLOIFOR','Libelle emploi',i);
                        If Combo.Value = 'RF' Then CreerRupture('PFO_RESPONSFOR','Responsable formation',i);
                        If Combo.Value = 'TY' Then CreerRupture('PST_NATUREFORM','Nature formation',i);
                        If Combo.Value = 'FO' Then CreerRupture('PFO_CODESTAGE','Formation',i);
                        If Combo.Value = 'S1' Then CreerRupture('PFO_TRAVAILN1',VH_Paie.PGLibelleOrgStat1,i);
                        If Combo.Value = 'S2' Then CreerRupture('PFO_TRAVAILN2',VH_Paie.PGLibelleOrgStat2,i);
                        If Combo.Value = 'S3' Then CreerRupture('PFO_TRAVAILN3',VH_Paie.PGLibelleOrgStat3,i);
                        If Combo.Value = 'S4' Then CreerRupture('PFO_TRAVAILN4',VH_Paie.PGLibelleOrgStat4,i);
                        If Combo.Value = 'ST' Then CreerRupture('PFO_CODESTAT',VH_Paie.PGLibCodeStat,i);
                end
                Else SetControlText('XX_RUPTURE'+IntToStr(i),'');
        end;
        SQL := 'SELECT COUNT(PFO_SALARIE) AS NBSTAGIAIRES,SUM(PFO_NBREHEURE) AS NBHEURES,'+GroupBy+
        ' FROM FORMATIONS LEFT JOIN STAGE ON PST_CODESTAGE=PFO_CODESTAGE AND PST_MILLESIME=PFO_MILLESIME '+Where+' GROUP BY '+GroupBy+' '+OrderBy;
        TFQRS1(Ecran).WhereSQL := SQL;
end;

procedure TOF_PGFORSUIVIGESTCOLL.CreerRupture(Champ,Tablette : String;Num : Integer);
begin
        SetControlText('XX_RUPTURE'+IntToStr(Num),Champ);
        SetControlText('CHAMP'+IntToStr(Num),Champ);
        SetControlText('TABLETTE'+IntToStr(Num),Tablette);
        If OrderBy <> '' Then OrderBy  := OrderBy+','+Champ
        Else OrderBy := 'ORDER BY '+Champ;
        If GroupBy = '' Then GroupBy := Champ
        Else GroupBy := GroupBy+','+Champ;
end;

{TOF_PGEDITFORSUIVIINV}
procedure TOF_PGEDITFORSUIVIINV.OnArgument (S : String ) ;
var i,Num : Integer;
    Combo : THValComboBox;
    Edit : THEdit;
    Arg ,Millesime: String;
    DD,DF : TDateTime;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        SetControlCaption('LIBSTAGE','');
        SetControlCaption('LIBLIEU','');
        Millesime := RendMillesimeRealise(DD,DF);
        SetControlText('PSS_DATEDEBUT',DateToStr(DD));
        SetControlText('PSS_DATEFIN',DateToStr(DF));
        Edit := THEdit(GetControl('PSS_DATEDEBUT'));
        If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('PSS_DATEFIN'));
        If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisClick;
        
        //PT5 - Début
        If PGBundleHierarchie Then //PT6
        Begin
	        Edit := THEdit(getControl('PFO_SALARIE'));
	        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	        Edit := THEdit(getControl('PFO_SALARIE_'));
	        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	        Edit := THEdit(getControl('PFO_RESPONSFOR'));
	        If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick;
	    End;
        //PT5 - Fin
        
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PSS_FORMATION'+IntToStr(Num)),GetControl ('TPSS_FORMATION'+IntToStr(Num)));
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
        end;
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('THVRUPTURE'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                        Combo.Items.Add('<<Aucun>>');
                        Combo.Values.Add('');
                        If Arg <> 'FISCALETFINANCIER' then    //PT3
                        begin
                                Combo.Items.Add('Type de financement');
                                Combo.Values.Add('PIS_INVESTFORM');
                        end;
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('PSS_CODESTAGE');
                        Combo.Items.Add('Etablissement');
                        Combo.Values.Add('PFO_ETABLISSEMENT');
                        Combo.Items.Add('Responsable');
                        Combo.Values.Add('PFO_RESPONSFOR');
                        Combo.Items.Add('Salarié');
                        Combo.Values.Add('PFO_SALARIE');
                        Combo.Items.Add('Nature (Interne/Externe)');
                        Combo.Values.Add('PSS_NATUREFORM');
                        Combo.Items.Add('Lieu de formation');
                        Combo.Values.Add('PSS_LIEUFORM');
                        For Num  :=  1 to VH_Paie.NBFormationLibre do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PSS_FORMATION1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PSS_FORMATION2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PSS_FORMATION3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PSS_FORMATION4');end;
                                If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PSS_FORMATION5');end;
                                If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PSS_FORMATION6');end;
                                If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PSS_FORMATION7');end;
                                If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PSS_FORMATION8');end;
                        end;
                        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('PFO_TRAVAILN1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('PFO_TRAVAILN2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('PFO_TRAVAILN3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('PFO_TRAVAILN4');end;
                        end;
                end;
        end;
        SetControlProperty('THVRUPTURE2','Value','');
        SetControlProperty('THVRUPTURE3','Value','');
        SetControlProperty('THVRUPTURE4','Value','');
        If Arg = 'FISCALETFINANCIER' then
        begin
                Ecran.Caption := 'Suivi fiscal et financier';
                UpdateCaption(TFQRS1(Ecran));
//                TFQRS1(Ecran).BValider.OnClick := EtatAvecTob; //PT5
                SetControlText('TYPECALC','FISCAL');
                SetControlVisible('PIS_INVESTFORM',False);
                SetControlVisible('TPIS_INVESTFORM',False);
                SetControlProperty('THVRUPTURE1','Value','PSS_CODESTAGE');
                TFQRS1(Ecran).CodeEtat := 'PFF';
                TFQRS1(Ecran).FCodeEtat := 'PFF';
        end
        else
        begin
                SetControlVisible('TYPECALC',False);
                SetControlVisible('TTYPECALC',False);
                SetControlProperty('THVRUPTURE1','Value','PIS_INVESTFORM');
                TFQRS1(Ecran).CodeEtat := 'PIV';
                TFQRS1(Ecran).FCodeEtat := 'PIV';
        end;
        {$IFDEF EMANAGER}
        SetControlText('PFO_RESPONSFOR',V_PGI.UserSalarie);
        SetControlVisible('PFO_RESPONSFOR',False);
        SetControlVisible('TPFO_RESPONSFOR',False);
        {$ENDIF}

		//PT5 - Début        
        If PGBundleInscFormation Or PGBundleCatalogue Then //PT6
        Begin
        	SetControlText('XX_WHEREPREDEF', DossiersAInterroger('',V_PGI.NoDossier,'PSS',True,False));
        	SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
        End;
        //PT5 - Fin
end;

procedure TOF_PGEDITFORSUIVIINV.OnUpdate;
var SQL,Where,OrderBy,Champ,GroupBy,StSelect : String;
    Pages  : TPageControl;
    i : Integer;
begin
  Inherited ;
  		//PT5 -  Début
  		If Assigned(TobEtat) Then FreeAndNil(TobEtat);
		TobEtat := TOB.Create('~LesDonnees', Nil, -1);
		//PT5 - Fin
		
		//PT5
        If (GetControlText('TYPECALC') = 'FISCAL') Or (GetControlText('TYPECALC') = 'FINANCIER') then
        begin
			EtatAvecTob;
           	TFQRS1(Ecran).LaTob := TobEtat; //PT5
        End
        Else
        Begin
	        Pages := TPageControl(GetControl('Pages'));
	        Where := RecupWhereCritere(Pages);
	        OrderBy := '';
	        GroupBy := 'GROUP BY PIS_CODESTAGE,PIS_ORDRE,PIS_TAUXPEDAG,'+
	        'PIS_TAUXSALFRAIS,PIS_TAUXSALAIRE,PSS_DATEDEBUT,PSS_DATEFIN';
	        StSelect := '';
	        For i := 1 to 4 do
	        begin
	                Champ := '';
	                Champ := GetControlText('THVRUPTURE'+IntToStr(i));
	                SetControlText('XX_RUPTURE'+IntToStr(i),Champ);
	                If Champ <> '' Then
	                begin
	                        If OrderBy <> '' Then OrderBy := OrderBy+','+Champ
	                        Else OrderBy := 'ORDER BY '+Champ;
	                        GroupBy := GroupBy+','+Champ;
	                        StSelect := StSelect + ','+ Champ;
	                end;
	        end;
	        If Where <> '' Then Where := ' '+Where+' '
	        else Where := ' ';
	        SQL := 'SELECT PIS_CODESTAGE,PIS_ORDRE,PIS_TAUXPEDAG,'+
	        'PIS_TAUXSALFRAIS,PIS_TAUXSALAIRE,PSS_DATEDEBUT,PSS_DATEFIN'+
	        ',SUM(PFO_AUTRECOUT) AS C1,SUM(PFO_COUTREELSAL) AS C2,SUM(PFO_FRAISREEL) AS C3'+StSelect+
	        ' FROM INVESTSESSION LEFT JOIN SESSIONSTAGE '+
	        'ON PIS_CODESTAGE=PSS_CODESTAGE AND PIS_ORDRE=PSS_ORDRE AND PIS_MILLESIME=PSS_MILLESIME '+
	        'LEFT JOIN FORMATIONS ON PFO_CODESTAGE=PIS_CODESTAGE AND PFO_ORDRE=PIS_ORDRE AND PFO_MILLESIME=PIS_MILLESIME '+
	        'AND PIS_ORDRE<>-1 ';  //DB2
        	TFQRS1(Ecran).WhereSQL := SQL+Where+GroupBy+' '+OrderBy; //PT5
        End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT5
Modifié le ... :   /  /
Description .. : Libération des éléments à la fermeture de l'état
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITFORSUIVIINV.OnClose ;
Begin
	If Assigned(TobEtat) Then FreeAndNil(TobEtat);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT5
Modifié le ... :   /  /
Description .. : Clic sur elipsis Salarie
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITFORSUIVIINV.SalarieElipsisClick(Sender : TObject);
begin
    //If PGBundleInscFormation Then
    	ElipsisSalarieMultidos (Sender)
    //Else
//    	Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT5
Modifié le ... :   /  /
Description .. : Clic sur elipsis responsable
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITFORSUIVIINV.RespElipsisClick(Sender : TObject);
begin
	ElipsisResponsableMultidos (Sender);
end;

procedure TOF_PGEDITFORSUIVIINV.EtatAvecTob; // (Sender : TObject); //PT5
var {TobEtat,}T,TobAnim,TA,TFRaisAnim,TF : Tob; //PT5
    Q : TQuery;
    CodeStage,Millesime,Where,OrderBy,Champ,GroupBy,StSelect,rupt1,Rupt2,Rupt3,Rupt4 : String;
    Tablette1,Tablette2,Tablette3,Tablette4 : String;
    Pages  : TPageControl;
    i : Integer;
    SalaireAnim,FraisAnim,CtPedag : Double;
    Session,NbHStagiaire,NbHStagiaireTot : Integer;
begin
        Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        OrderBy := '';
        Rupt1 := GetControlText('THVRUPTURE1');
        Rupt2 := GetControlText('THVRUPTURE2');
        Rupt3 := GetControlText('THVRUPTURE3');
        Rupt4 := GetControlText('THVRUPTURE4');
        Tablette1 := ChampRuptureForm(Rupt1);
        Tablette2 := ChampRuptureForm(Rupt2);
        Tablette3 := ChampRuptureForm(Rupt3);
        Tablette4 := ChampRuptureForm(Rupt4);
        GroupBy := 'GROUP BY PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_CODESTAGE,PSS_MILLESIME,PSS_ORDRE';
        StSelect := '';
        For i := 1 to 4 do
        begin
                Champ := '';
                Champ := GetControlText('THVRUPTURE'+IntToStr(i));
                SetControlText('XX_RUPTURE'+IntToStr(i),Champ);
                If Champ <> '' Then
                begin
                        If OrderBy <> '' Then OrderBy := OrderBy+','+Champ
                        Else OrderBy := 'ORDER BY '+Champ;
                        If Champ <> 'PSS_CODESTAGE' then
                        begin
                                GroupBy := GroupBy+','+Champ;
                                StSelect := StSelect + ','+ Champ;
                        end;
                end;
        end;
        If Where <> '' Then Where := ' '+Where+' '
        else Where := ' ';
        
        If GetControlText('TYPECALC') = 'FISCAL' then
        begin
                Q := OpenSQL('SELECT PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_CODESTAGE,PSS_MILLESIME,PSS_ORDRE'+
                ',SUM(PFO_AUTRECOUT) CPEDAGOGIQUE,SUM(PFO_COUTREELSAL) SALAIRES,SUM(PFO_FRAISREEL) FRAIS'+StSelect+ //PT5
                ' FROM SESSIONSTAGE '+
                'LEFT JOIN FORMATIONS ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME AND PFO_EFFECTUE="X" '+
                Where+GroupBy+' '+OrderBy,True);
//                TobEtat := Tob.Create('LesFrais',Nil,-1); //PT5
                TobEtat.LoadDetailDB('LesFrais','','',Q,False);
                Ferme(Q);
                For i := 0 to TobEtat.Detail.Count-1 do
                begin
                        T:= TobEtat.Detail[i];
                        T.AddChampSup('RUPTURE1',False);
                        T.AddChampSup('RUPTURE2',False);
                        T.AddChampSup('RUPTURE3',False);
                        T.AddChampSup('RUPTURE4',False);
                        T.AddChampSup('LIBRUPTURE1',False);
                        T.AddChampSup('LIBRUPTURE2',False);
                        T.AddChampSup('LIBRUPTURE3',False);
                        T.AddChampSup('LIBRUPTURE4',False);
                        T.AddChampSup('CPEDAGOGIQUE',False);
                        T.AddChampSup('SALAIRES',False);
                        T.AddChampSup('FRAIS',False);
                        If Rupt1 <> '' then
                        begin
                                T.PutValue('RUPTURE1',TobEtat.Detail[i].GetValue(Rupt1));
                                T.PutValue('LIBRUPTURE1',RechDom(Tablette1,TobEtat.Detail[i].GetValue(Rupt1),False));
                        end
                        else begin T.PutValue('RUPTURE1',''); T.PutValue('LIBRUPTURE1',''); end;
                        If Rupt2 <> '' then
                        begin
                                T.PutValue('RUPTURE2',TobEtat.Detail[i].GetValue(Rupt2));
                                T.PutValue('LIBRUPTURE2',RechDom(Tablette2,TobEtat.Detail[i].GetValue(Rupt2),False));
                        end
                        else begin T.PutValue('RUPTURE2',''); T.PutValue('LIBRUPTURE2',''); end;
                        If Rupt3 <> '' then
                        begin
                                T.PutValue('RUPTURE3',TobEtat.Detail[i].GetValue(Rupt3));
                                T.PutValue('LIBRUPTURE3',RechDom(Tablette3,TobEtat.Detail[i].GetValue(Rupt3),False));
                        end
                        else begin T.PutValue('RUPTURE3',''); T.PutValue('LIBRUPTURE3',''); end;
                        If Rupt4 <> '' then
                        begin
                                T.PutValue('RUPTURE4',TobEtat.Detail[i].GetValue(Rupt4));
                                T.PutValue('LIBRUPTURE4',RechDom(Tablette4,TobEtat.Detail[i].GetValue(Rupt4),False));
                        end
                        else begin T.PutValue('RUPTURE4',''); T.PutValue('LIBRUPTURE4',''); end;
                        //PT5
                        // Cout pédagogique = Coût pédagogique + salaire formateurs + frais formateurs
                        //T.PutValue('CPEDAGOGIQUE',TobEtat.Detail[i].GetValue('CTPEDAG'));
                        // Salaire = Salaires stagiaires
                        //T.PutValue('SALAIRES',TobEtat.Detail[i].GetValue('CTSALAIRE'));
                        //Frais = Frais Stagiaire
                        //T.PutValue('FRAIS',TobEtat.Detail[i].GetValue('CTFRAIS'));
                end;
                TobEtat.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;RUPTURE4');
                //PT5
                //If getCheckBoxState('FLISTE') = CbChecked then
              	//LanceEtatTOB('E','PFO','PFF',TobEtat,True,True,False,Pages,'','',False)
                //else
               	//LanceEtatTOB('E','PFO','PFF',TobEtat,True,False,False,Pages,'','',False);
                //TobEtat.Free;
        end;
        If GetControlText('TYPECALC') = 'FINANCIER' then
        begin
                Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) NBHSAL,SUM(PFO_COUTREELSAL) CTSALAIRE,SUM(PFO_FRAISREEL) CTFRAIS,PSS_LIBELLE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_MILLESIME,PSS_CODESTAGE,PSS_ORDRE'+StSelect+' FROM SESSIONSTAGE '+
                'LEFT JOIN FORMATIONS ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME AND PFO_EFFECTUE="X" '+
                Where+GroupBy+' '+OrderBy,True);
//                TobEtat := Tob.Create('LesCouts',Nil,-1); //PT5
                TobEtat.LoadDetailDB('LesCouts','','',Q,False);
                Ferme(Q);
                
                Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_COUTPEDAG,PSS_NATUREFORM,SUM (PAN_SALAIREANIM) SALAIREANIM FROM SESSIONSTAGE '+
                'LEFT JOIN SESSIONANIMAT ON PSS_CODESTAGE=PAN_CODESTAGE AND PSS_ORDRE=PAN_ORDRE AND PSS_MILLESIME=PAN_MILLESIME '+
                'WHERE PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('PSS_DATEDEBUT')))+'" AND PSS_DATEFIN<="'+UsdateTime(StRtoDate(GetControlText('PSS_DATEFIN')))+'" '+
                'GROUP BY PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_COUTPEDAG,PSS_NATUREFORM',True);
                TobAnim := Tob.Create('LesAnimateurs',Nil,-1);
                TobAnim.LoadDetailDB('LesAnimateurs','','',Q,False);
                Ferme(Q);
                
                Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_COUTPEDAG,PSS_NATUREFORM,SUM(PFS_MONTANT) FRAISANIM FROM SESSIONSTAGE '+
                'LEFT JOIN SESSIONANIMAT ON PSS_CODESTAGE=PAN_CODESTAGE AND PSS_ORDRE=PAN_ORDRE AND PSS_MILLESIME=PAN_MILLESIME '+
                'LEFT JOIN FRAISSALFORM ON PAN_SALARIE=PFS_SALARIE AND PAN_CODESTAGE=PFS_CODESTAGE AND PAN_ORDRE=PFS_ORDRE AND PAN_MILLESIME=PFS_MILLESIME '+
                'WHERE PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('PSS_DATEDEBUT')))+'" AND PSS_DATEFIN<="'+UsdateTime(StRtoDate(GetControlText('PSS_DATEFIN')))+'" '+
                'GROUP BY PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_COUTPEDAG,PSS_NATUREFORM',True);
                TFRaisAnim := Tob.Create('LesAnimateurs',Nil,-1);
                TFRaisAnim.LoadDetailDB('LesAnimateurs','','',Q,False);
                Ferme(Q);
                
                For i := 0 to TobEtat.Detail.Count-1 do
                begin
                        T:= TobEtat.Detail[i];
                        T.AddChampSup('RUPTURE1',False);
                        T.AddChampSup('RUPTURE2',False);
                        T.AddChampSup('RUPTURE3',False);
                        T.AddChampSup('RUPTURE4',False);
                        T.AddChampSup('LIBRUPTURE1',False);
                        T.AddChampSup('LIBRUPTURE2',False);
                        T.AddChampSup('LIBRUPTURE3',False);
                        T.AddChampSup('LIBRUPTURE4',False);
                        T.AddChampSup('CPEDAGOGIQUE',False);
                        T.AddChampSup('SALAIRES',False);
                        T.AddChampSup('FRAIS',False);
                        If Rupt1 <> '' then
                        begin
                                T.PutValue('RUPTURE1',TobEtat.Detail[i].GetValue(Rupt1));
                                T.PutValue('LIBRUPTURE1',RechDom(Tablette1,TobEtat.Detail[i].GetValue(Rupt1),False));
                        end
                        else begin T.PutValue('RUPTURE1',''); T.PutValue('LIBRUPTURE1',''); end;
                        If Rupt2 <> '' then
                        begin
                                T.PutValue('RUPTURE2',TobEtat.Detail[i].GetValue(Rupt2));
                                T.PutValue('LIBRUPTURE2',RechDom(Tablette2,TobEtat.Detail[i].GetValue(Rupt2),False));
                        end
                        else begin T.PutValue('RUPTURE2',''); T.PutValue('LIBRUPTURE2',''); end;
                        If Rupt3 <> '' then
                        begin
                                T.PutValue('RUPTURE3',TobEtat.Detail[i].GetValue(Rupt3));
                                T.PutValue('LIBRUPTURE3',RechDom(Tablette3,TobEtat.Detail[i].GetValue(Rupt3),False));
                        end
                        else begin T.PutValue('RUPTURE3',''); T.PutValue('LIBRUPTURE3',''); end;
                        If Rupt4 <> '' then
                        begin
                                T.PutValue('RUPTURE4',TobEtat.Detail[i].GetValue(Rupt4));
                                T.PutValue('LIBRUPTURE4',RechDom(Tablette4,TobEtat.Detail[i].GetValue(Rupt4),False));
                        end
                        else begin T.PutValue('RUPTURE4',''); T.PutValue('LIBRUPTURE4',''); end;
                        SalaireAnim := 0;
                        FraisAnim := 0;
                        CtPedag := 0;
                        if T.GetValue('NBHSAL') <> Null then NbHStagiaire := T.GetValue('NBHSAL')
                        else NbHStagiaire := 0;
                        CodeStage := T.GetValue('PSS_CODESTAGE');
                        Session := T.GetValue('PSS_ORDRE');
                        Millesime := T.GetValue('PSS_MILLESIME');
                        Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) NBHSALTOT FROM FORMATIONS WHERE PFO_CODESTAGE="'+CodeStage+'"'+
                        ' AND PFO_ORDRE='+IntToStr(Session)+' AND PFO_MILLESIME="'+Millesime+'"',True);
                        If Not Q.Eof then NbHStagiaireTot := Q.FindField('NBHSALTOT').AsInteger
                        else NbHStagiaireTot := 1;
                        Ferme(Q);
                        TA := TobAnim.FindFirst(['PSS_CODESTAGE','PSS_ORDRE','PSS_MILLESIME'],[CodeStage,Session,Millesime],False);
                        If TA <> Nil then
                        begin

                                If TA.GetValue('PSS_NATUREFORM') = '002' then CtPedag := TA.GetValue('PSS_COUTPEDAG');
                                If (TA.GetValue('SALAIREANIM') <> NULL) and isnumeric(TA.GetValue('SALAIREANIM')) then SalaireAnim := TA.GetValue('SALAIREANIM');  // PT2
                        end;
                        TF := TFRaisAnim.FindFirst(['PSS_CODESTAGE','PSS_ORDRE','PSS_MILLESIME'],[CodeStage,Session,Millesime],False);
                        If TF <> Nil then
                        begin
                                If (TF.GetValue('FRAISANIM') <> Null) and isnumeric(TF.GetValue('FRAISANIM'))then FraisAnim := TF.GetValue('FRAISANIM');  // PT2
                        end;
                        If NbHStagiaireTot <> 0 then
                        begin
                                SalaireAnim := (SalaireANim*NbHStagiaire)/NbHStagiaireTot;
                                FraisAnim := (FraisAnim*NbHStagiaire)/NbHStagiaireTot;
                                CtPedag := (CtPedag*NbHStagiaire)/NbHStagiaireTot;
                        end
                        else
                        begin
                                SalaireAnim := 0;
                                FraisAnim := 0;
                                CtPedag := 0;
                        end;
                        // Cout pédagogique = Coût pédagogique externe       *** pas de salaire et frais formateur
                        T.PutValue('CPEDAGOGIQUE',CtPedag);
                        // Salaire = Salaire Stagiaire + salaire animateur
                        T.PutValue('SALAIRES',TobEtat.Detail[i].GetValue('CTSALAIRE')+SalaireAnim);
                        // Frais = Frais stagiaire + frais animateur
                        T.PutValue('FRAIS',TobEtat.Detail[i].GetValue('CTFRAIS') + FraisANim);
                end;
                TobAnim.Free;
                TFRaisAnim.Free;
                TobEtat.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;RUPTURE4');
                //PT5
                //If GetCheckBoxState('FLISTE') = CbChecked then
                //LanceEtatTOB('E','PFO','PFF',TobEtat,True,True,False,Pages,'','',False)
                //else
               	//LanceEtatTOB('E','PFO','PFF',TobEtat,True,False,False,Pages,'','',False);
                //TobEtat.Free;
        end;
end;

Function TOF_PGEDITFORSUIVIINV.ChampRuptureForm(ValeurRupt : String) : String;
var Tablette : String;
    j : Integer;
begin
        If ValeurRupt = '' then
        begin
                result := '';
                Exit;
        end;
        Tablette := '';
        If ValeurRupt = 'PSS_CODESTAGE' then Tablette := 'PGSTAGEFORM'
        else If ValeurRupt = 'PFO_SALARIE' then 
        Begin
        	If PGBundleHierarchie Then //PT5
        		Tablette := 'PGSALARIEINT'
        	Else
        		Tablette := 'PGSALARIE';
        End
        else If ValeurRupt = 'PFO_ETABLISSEMENT' then Tablette := 'TTETABLISSEMENT'
        else If ValeurRupt = 'PFO_LIBEMPLOIFOR' then Tablette := 'PGLIBEMPLOI'
        else If ValeurRupt = 'PSS_NATUREFORM' then Tablette := 'PGNATUREFORM'
        else If ValeurRupt = 'PSS_LIEUFORM' then Tablette := 'PGLIEUFORMATION';
        If Tablette = '' then
        begin
                For j := 1 to 8 do
                begin
                        If ValeurRupt = 'PFO_FORMATION'+IntToStr(j) then Tablette := 'PGFORMATION'+IntToStr(j);
                end;
        end;
        If Tablette ='' then
        begin
                For j := 1 to 4 do
                begin
                        If ValeurRupt = 'PFO_TRAVAILN'+IntToStr(j) then Tablette := 'PGTRAVAILN'+IntToStr(j);
                end;
        end;
        Result := Tablette;
end;

procedure TOF_PGEDITFORSUIVIINV.DateElipsisClick (Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

{TOF_PGFORCOMPCOLL}
procedure TOF_PGFORCOMPCOLL.OnArgument(S : String);
var Num,i : Integer;
    Combo : THValComboBox;
    Millesime : String;
    DD,DF : TDateTime;
    BEtat : TToolBarButton97;
    Arg,Min,Max : String;
     Edit : THEdit;
begin
  Inherited ;
        Arg := S;
        If Arg = 'CWAS' then
        begin
                SetControlText('PFO_RESPONSFOR',V_PGI.UserSalarie);
                SetControlVisible('PFO_RESPONSFOR',False);
                SetControlVisible('TPFO_RESPPONSFOR',False);
                Edit := THEdit(GetControl('PFO_SALARIE'));
                If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
                Edit := THEdit(GetControl('PFO_SALARIE_'));
                If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        end
        else
        begin
                RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
                SetControlText('PFO_SALARIE',Min);
                SetControlText('PFO_SALARIE_',Max);
                SetControlCaption('LIBRESP','');
        end;
        Millesime := RendMillesimeRealise(DD,DF);
        SetControltext('DATEDEBUT',DateToStr(DD));
        SetControltext('DATEFIN',DateToStr(DF));
        SetControlText('MILLESIME',Millesime);
        Betat  := TToolBarButton97(GetControl('BVALIDER'));
        If BEtat <> Nil then BEtat.OnClick  := BEtatClick;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PSS_FORMATION'+IntToStr(Num)),GetControl ('TPSS_FORMATION'+IntToStr(Num)));
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
        end;
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                        If i <> 1 then
                        begin
                                Combo.Items.Add('<<Aucun>>');
                                Combo.Values.Add('');
                        end;
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('PSS_CODESTAGE');
                        Combo.Items.Add('Etablissement');
                        Combo.Values.Add('PFO_ETABLISSEMENT');
                        Combo.Items.Add('Responsable');
                        Combo.Values.Add('PFO_RESPONSFOR');
                        Combo.Items.Add('Salarié');
                        Combo.Values.Add('PFO_SALARIE');
                        Combo.Items.Add('Nature (Interne/Externe)');
                        Combo.Values.Add('PSS_NATUREFORM');
                        Combo.Items.Add('Lieu de formation');
                        Combo.Values.Add('PSS_LIEUFORM');
                        For Num  :=  1 to VH_Paie.NBFormationLibre do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PSS_FORMATION1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PSS_FORMATION2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PSS_FORMATION3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PSS_FORMATION4');end;
                                If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PSS_FORMATION5');end;
                                If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PSS_FORMATION6');end;
                                If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PSS_FORMATION7');end;
                                If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PSS_FORMATION8');end;
                        end;
                        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('PSA_TRAVAILN1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('PSA_TRAVAILN2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('PSA_TRAVAILN3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('PSA_TRAVAILN4');end;
                        end;
                end;
        end;
        SetControlText('VRUPT1','PSS_CODESTAGE');
        SetControlText('VRUPT2','');
        SetControlText('VRUPT3','');
        SetControlText('VRUPT4','');
        Combo := THValComboBox(GetControl('MILLESIME'));
        If Combo <> Nil then Combo.OnChange := ChangeMillesime;
        SetControlCaption('LIBSTAGE','');
end;

procedure TOF_PGFORCOMPCOLL.SalarieElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        StOrder := 'PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT4
        LookupList(THEdit(Sender),'Liste des stages','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;

procedure TOF_PGFORCOMPCOLL.BEtatClick(Sender : TObject);
var TobNbInscBud,TobNbInscRea,TobBudget,TobRealise,TobEtat : Tob;
     TNbIB,TNbIR,T : Tob;
     Q : TQuery;
     DateDebut,DateFin : TDateTime;
     i,NbRupt : Integer;
     Salarie,Where,WherePFI,WherePFO,Millesime,CodeStage,StPages : String;
     Rupt1,Rupt2,Rupt3,Rupt4 : String;
     ChampRuptRea1,ChampRuptRea2,ChampRuptRea3,ChampRuptRea4 : String;
     ChampRuptBud1,ChampRuptBud2,ChampRuptBud3,ChampRuptBud4 : String;
     ValeurRupt1,ValeurRupt2,ValeurRupt3,ValeurRupt4 : String;
     Tablette1,Tablette2,Tablette3,Tablette4,Libelle : String;
     NbStaMax,NbInscSal,NbInsc : Integer;
     NbHeuresR,NbHeuresB,CoutPedag,CoutAnim,CoutUnitaire,CoutSal,Frais : Double;
     TotalRea,TotalBud,NbSession : Double;
     Pages : TPageControl;
     ListeChampRea,ListeChampBud : String;
begin
        Pages := TPageControl(GetControl('Pages'));
        Rupt1 := '';
        Rupt2 := '';
        Rupt3 := '';
        Rupt4 := '';
        NbInsc := 0;
        If GetControltext('VRUPT4') <> '' then
        begin
                Rupt4 := GetControltext('VRUPT4');
                Rupt3 := GetControltext('VRUPT3');
                Rupt2 := GetControltext('VRUPT2');
                Rupt1 := GetControltext('VRUPT1');
                SetControlText('XX_RUPTURE1','RUPTURE1');
                SetControlText('XX_RUPTURE2','RUPTURE2');
                SetControlText('XX_RUPTURE3','RUPTURE3');
                SetControlText('XX_RUPTURE4','RUPTURE4');
                NbRupt := 4;
                If GetControlText('VRUPT3') = '' then begin PGIBox('Le niveau de rupture 3 doit être renseigné',Ecran.Caption);exit;end;
                If GetControlText('VRUPT2') = '' then begin PGIBox('Le niveau de rupture 2 doit être renseigné',Ecran.Caption);exit;end;
        end
        Else
        begin
                If GetControlText('VRUPT3') <> '' then
                begin
                        Rupt3 := GetControltext('VRUPT3');
                        Rupt2 := GetControltext('VRUPT2');
                        Rupt1 := GetControltext('VRUPT1');
                        SetControlText('XX_RUPTURE1','RUPTURE1');
                        SetControlText('XX_RUPTURE2','RUPTURE2');
                        SetControlText('XX_RUPTURE3','RUPTURE3');
                        SetControlText('XX_RUPTURE4','');
                        NbRupt := 3;
                        If GetControlText('VRUPT2') = '' then begin PGIBox('Le niveau de rupture 2 doit être renseigné',Ecran.Caption);exit;end;
                end
                Else
                begin
                        If GetControlText('VRUPT2') <> '' then
                        begin
                                Rupt2 := GetControltext('VRUPT2');
                                Rupt1 := GetControltext('VRUPT1');
                                SetControlText('XX_RUPTURE1','RUPTURE1');
                                SetControlText('XX_RUPTURE2','RUPTURE2');
                                SetControlText('XX_RUPTURE3','');
                                SetControlText('XX_RUPTURE4','');
                                NbRupt := 2;
                        end
                        Else
                        begin
                                Rupt1 := GetControltext('VRUPT1');
                                SetControlText('XX_RUPTURE1','RUPTURE1');
                                SetControlText('XX_RUPTURE2','');
                                SetControlText('XX_RUPTURE3','');
                                SetControlText('XX_RUPTURE4','');
                                NbRupt := 1;
                        end;
                end;
        end;
        SetControlText('NBRUPT',IntToStr(NbRupt));
        ChampRuptRea1 := Rupt1;
        ChampRuptRea2 := Rupt2;
        ChampRuptRea3 := Rupt3;
        ChampRuptRea4 := Rupt4;
        ChampRuptBud1 := ConvertPrefixe(Rupt1,'PSS','PST');
        ChampRuptBud2 := ConvertPrefixe(Rupt2,'PSS','PST');
        ChampRuptBud3 := ConvertPrefixe(Rupt3,'PSS','PST');
        ChampRuptBud4 := ConvertPrefixe(Rupt4,'PSS','PST');
        ChampRuptBud1 := ConvertPrefixe(ChampRuptBud1,'PFO','PFI');
        ChampRuptBud2 := ConvertPrefixe(ChampRuptBud2,'PFO','PFI');
        ChampRuptBud3 := ConvertPrefixe(ChampRuptBud3,'PFO','PFI');
        ChampRuptBud4 := ConvertPrefixe(ChampRuptBud4,'PFO','PFI');
        If Rupt1 <> '' then ListeChampRea := ListeChampRea+','+Rupt1;
        If Rupt2 <> '' then ListeChampRea := ListeChampRea+','+Rupt2;
        If Rupt3 <> '' then ListeChampRea := ListeChampRea+','+Rupt3;
        If Rupt4 <> '' then ListeChampRea := ListeChampRea+','+Rupt4;
        If ChampRuptBud1 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud1;
        If ChampRuptBud2 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud2;
        If ChampRuptBud3 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud3;
        If ChampRuptBud4 <> '' then ListeChampBud := ListeChampBud+','+ChampRuptBud4;
        Tablette1 := ChampRuptureForm('TAB',Rupt1);
        Tablette2 := ChampRuptureForm('TAB',Rupt2);
        Tablette3 := ChampRuptureForm('TAB',Rupt3);
        Tablette4 := ChampRuptureForm('TAB',Rupt4);

        DateDebut := StrToDate(GetControlText('DATEDEBUT'));
        dateFin := StrToDate(GetControlText('DATEFIN'));

        Where := '';
        Where := RecupWhereCritere(Pages);
        WherePFI := ConvertPrefixe(Where,'PFO','PFI');
        WherePFI := ConvertPrefixe(WherePFI,'PSS','PST');
        WherePFO := Where;
        If WherePFO <> '' then WherePFO := WherePFO+' AND PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PFO_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'" AND PFO_EFFECTUE="X"'
        else WherePFO := 'WHERE PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'" AND PFO_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'" AND PFO_EFFECTUE="X"';
        If WherePFI <> '' then WHerePFI := WherePFI + 'AND PFI_ETATINSCFOR="VAL" AND PFI_MILLESIME="'+GetControlText('MILLESIME')+'"'
        Else WHerePFI := 'WHERE PFI_ETATINSCFOR="VAL" AND PFI_MILLESIME="'+GetControlText('MILLESIME')+'"';
        If (ChampRuptBud1 = 'PFI_SALARIE') or (ChampRuptBud2 = 'PFI_SALARIE') or (ChampRuptBud3 = 'PFI_SALARIE') or (ChampRuptBud4 = 'PFI_SALARIE') then
        Where := Where + ' AND PFI_SALARIE <> ""';
        Q := OpenSQL('SELECT PFI_MILLESIME,PFI_CODESTAGE,PFI_NBINSC,PFI_COUTREELSAL,PFI_FRAISFORFAIT,PFI_SALARIE,PFI_LIBELLE,PFI_LIBEMPLOIFOR,PFI_ETABLISSEMENT,PFI_RESPONSFOR'+
        ',PST_DUREESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX'+ListeChampBud+
        ' FROM INSCFORMATION LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME '+WherePFI,True);
        TobBudget := Tob.Create('Les inscriptions au budget',Nil,-1);
        TobBudget.LoadDetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM(PFI_NBINSC) AS NBINSC,PFI_MILLESIME,PFI_CODESTAGE FROM INSCFORMATION'+
        ' WHERE PFI_MILLESIME="'+GetcontrolText('MILLESIME')+'"'+
        ' GROUP BY PFI_CODESTAGE,PFI_MILLESIME',True);
        TobNbInscBud := Tob.Create('Nombre insc',Nil,-1);
        TobNbInscBud.LoadDetailDB('Nombre insc','','',Q,False);
        Ferme(Q);
        TobEtat := Tob.Create('Les salariés pour édition',Nil,-1);
        For i := 0 to TobBudget.Detail.Count-1 do
        begin
                Salarie := TobBudget.Detail[i].GetValue('PFI_SALARIE');
                CodeStage := TobBudget.Detail[i].GetValue('PFI_CODESTAGE');
                Millesime := TobBudget.Detail[i].GetValue('PFI_MILLESIME');
                TNbIB := TobNbInscBud.FindFirst(['PFI_CODESTAGE','PFI_MILLESIME'],[CodeStage,Millesime],False);
                If TNbIB <> Nil then NbInsc := TNbIB.GetValue('NBINSC')
                Else Exit;
                NbStaMax := TobBudget.Detail[i].GetValue('PST_NBSTAMAX');
                NbSession := NombreDeSessionFormationAPrevoir(NbStaMax,NbInsc);
                NbInscSal := TobBudget.Detail[i].GetValue('PFI_NBINSC');
                NbHeuresB := TobBudget.Detail[i].GetValue('PST_DUREESTAGE');
                NbHeuresB := NbHeuresB * NbInscSal;
                CoutPedag := TobBudget.Detail[i].GetValue('PST_COUTBUDGETE');
                CoutPedag := CoutPedag*NbSession*NbInscSal;
                If NbInsc > 0 then CoutPedag := arrondi(CoutPedag/NbInsc,2)
                else CoutPedag :=0;
                CoutAnim := TobBudget.Detail[i].GetValue('PST_COUTSALAIR');
                CoutAnim := CoutAnim*NbSession*NbInscSal;
                CoutAnim := arrondi(CoutAnim/NbInsc,2);
                CoutUnitaire := TobBudget.Detail[i].GetValue('PST_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire*NbInscSal;
                CoutSal := TobBudget.Detail[i].GetValue('PFI_COUTREELSAL');
                Frais := TobBudget.Detail[i].GetValue('PFI_FRAISFORFAIT');
                TotalBud := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                ValeurRupt1 := '';
                ValeurRupt2 := '';
                ValeurRupt3 := '';
                ValeurRupt4 := '';
                If NbRupt > 3 then ValeurRupt4 := TobBudget.Detail[i].GetValue(ChampRuptBud4);
                If NbRupt > 2 then ValeurRupt3 := TobBudget.Detail[i].GetValue(ChampRuptBud3);
                If NbRupt > 1 then ValeurRupt2 := TobBudget.Detail[i].GetValue(ChampRuptBud2);
                ValeurRupt1 := TobBudget.Detail[i].GetValue(ChampRuptBud1);
                If NbRupt=4 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt4,ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=3 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE2','RUPTURE1'],[ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=2 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE1'],[ValeurRupt2,ValeurRupt1],False)
                Else T := TobEtat.FindFirst(['RUPTURE4'],[ValeurRupt1],False);
                If T <> Nil then
                begin
                        T.PutValue('HEURESNR',T.GetValue('HEURESNR')+NbHeuresB);
                        T.PutValue('COUTSNR',T.GetValue('COUTSNR')+TotalBud);
                end
                Else
                begin
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSup('RUPTURE1',False);T.AddChampSup('RUPTURE2',False);
                        T.AddChampSup('RUPTURE3',False);T.AddChampSup('RUPTURE4',False);
                        T.AddChampSup('LIBRUPTURE1',False);T.AddChampSup('LIBRUPTURE2',False);
                        T.AddChampSup('LIBRUPTURE3',False);T.AddChampSup('LIBRUPTURE4',False);
                        T.AddChampSup('HEURESR',False);T.AddChampSup('COUTSR',False);
                        T.AddChampSup('HEURESNR',False);T.AddChampSup('COUTSNR',False);
                        T.AddChampSup('HEURESNP',False);T.AddChampSup('COUTSNP',False);
                        T.PutValue('HEURESNR',NbHeuresB);
                        T.PutValue('COUTSNR',TotalBud);
                        If NbRupt > 3 then T.PutValue('RUPTURE4',TobBudget.Detail[i].GetValue(ChampRuptBud4));
                        If NbRupt > 2 then T.PutValue('RUPTURE3',TobBudget.Detail[i].GetValue(ChampRuptBud3));
                        If NbRupt > 1 then T.PutValue('RUPTURE2',TobBudget.Detail[i].GetValue(ChampRuptBud2));
                        T.PutValue('RUPTURE1',TobBudget.Detail[i].GetValue(ChampRuptBud1));
                        T.PutValue('HEURESR',0);T.PutValue('COUTSR',0);
                        T.PutValue('HEURESNP',0);T.PutValue('COUTSNP',0);
                end;
        end;
        TobNbInscBud.Free;
        TobBudget.Free;
        Q := OpenSQL('SELECT PSS_COUTPEDAG,PSS_COUTSALAIR,PSS_COUTUNITAIRE,PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE,PFO_MILLESIME,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_RESPONSFOR'+
        ',PFO_NBREHEURE,PFO_FRAISREEL,PFO_COUTREELSAL,PFO_NBREHEURE'+ListeChampRea+
        ' FROM FORMATIONS LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME '+WherePFO,True);
        TobRealise := Tob.Create('Les inscriptions au formations',Nil,-1);
        TobRealise.LoadDetailDB('FORMATIONS','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) AS NBSAL,PFO_ORDRE,PFO_MILLESIME,PFO_CODESTAGE FROM FORMATIONS WHERE '+
        'PFO_DATEDEBUT>="'+UsDatEtime(DateDebut)+'" AND PFO_DATEDEBUT<="'+UsDateTime(Datefin)+'" AND PFO_EFFECTUE="X"'+
        ' GROUP BY PFO_ORDRE,PFO_CODESTAGE,PFO_MILLESIME',True);
        TobNbInscRea := Tob.Create('Les Sessions',Nil,-1);
        TobNbInscRea.LoadDetailDB('SESSIONSTAGE','','',Q,False);
        Ferme(Q);
        For i := 0 to TobRealise.detail.Count-1 do
        begin
                Salarie := TobRealise.Detail[i].GetValue('PFO_SALARIE');
                CodeStage := TobRealise.Detail[i].GetValue('PFO_CODESTAGE');
                Millesime := TobRealise.Detail[i].GetValue('PFO_MILLESIME');
                NbHeuresR := TobRealise.Detail[i].GetValue('PFO_NBREHEURE');
                TNbIR := TobNbInscRea.FindFirst(['PFO_CODESTAGE','PFO_MILLESIME','PFO_ORDRE'],[CodeStage,Millesime,TobRealise.Detail[i].GetValue('PFO_ORDRE')],False);
                If TNbIR <> Nil then NbInsc := TNbIR.GetValue('NBSAL')
                Else Exit;
                CoutPedag := TobRealise.Detail[i].GetValue('PSS_COUTPEDAG');
                CoutPedag := arrondi(NbHeuresR*(CoutPedag/NbInsc),2);
                CoutAnim := TobRealise.Detail[i].GetValue('PSS_COUTSALAIR');
                CoutAnim := arrondi(NbHeuresR*(CoutAnim/NbInsc),2);
                CoutUnitaire := TobRealise.Detail[i].GetValue('PSS_COUTUNITAIRE');
                CoutUnitaire := CoutUnitaire;
                CoutSal := TobRealise.Detail[i].GetValue('PFO_COUTREELSAL');
                Frais := TobRealise.Detail[i].GetValue('PFO_FRAISREEL');
                TotalRea := CoutPedag+CoutAnim+CoutUnitaire+CoutSal+Frais;
                ValeurRupt1 := '';
                ValeurRupt2 := '';
                ValeurRupt3 := '';
                ValeurRupt4 := '';
                If NbRupt > 3 then ValeurRupt4 := TobRealise.Detail[i].GetValue(ChampRuptRea4);
                If NbRupt > 2 then ValeurRupt3 := TobRealise.Detail[i].GetValue(ChampRuptRea3);
                If NbRupt > 1 then ValeurRupt2 := TobRealise.Detail[i].GetValue(ChampRuptRea2);
                ValeurRupt1 := TobRealise.Detail[i].GetValue(ChampRuptRea1);
                If NbRupt=4 then T := TobEtat.FindFirst(['RUPTURE4','RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt4,ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=3 then T := TobEtat.FindFirst(['RUPTURE3','RUPTURE2','RUPTURE1'],[ValeurRupt3,ValeurRupt2,ValeurRupt1],False)
                Else If NbRupt=2 then T := TobEtat.FindFirst(['RUPTURE2','RUPTURE1'],[ValeurRupt2,ValeurRupt1],False)
                Else T := TobEtat.FindFirst(['RUPTURE1'],[ValeurRupt1],False);
                If T <> Nil then
                begin
                        NbHeuresB := T.GetValue('HEURESNR');
                        TotalBud := T.GetValue('COUTSNR');
                        If NbHeuresR>0 then
                        begin
                                if NbHeuresB>=NbHeuresR then
                                begin
                                        T.PutValue('HEURESNR',NbHeuresB-NbHeuresR);
                                        T.PutValue('HEURESR',T.GetValue('HEURESR')+NbHeuresR);
                                end
                                Else
                                begin
                                        T.PutValue('HEURESNP',T.GetValue('HEURESNP') + NbHeuresR - T.GetValue('HEURESNR'));
                                        T.PutValue('HEURESNR',0);
                                        T.PutValue('HEURESR',T.GetValue('HEURESR')+NbHeuresB);
                                end;
                        end;
                        If TotalRea>0 then
                        begin
                                if TotalBud>=TotalRea then
                                begin
                                        T.PutValue('COUTSNR',TotalBud-TotalRea);
                                        T.PutValue('COUTSR',T.GetValue('COUTSR')+TotalRea);
                                end
                                Else
                                begin
                                        T.PutValue('COUTSNP',T.GetValue('COUTSNP') + TotalRea - T.GetValue('COUTSNR'));
                                        T.PutValue('COUTSNR',0);
                                        T.PutValue('COUTSR',T.GetValue('COUTSR')+TotalBud);
                                end;
                        end;
                end
                Else
                begin
                        T := Tob.Create('Les filles salariés',TobEtat,-1);
                        T.AddChampSup('RUPTURE1',False);T.AddChampSup('RUPTURE2',False);
                        T.AddChampSup('RUPTURE3',False);T.AddChampSup('RUPTURE4',False);
                        T.AddChampSup('LIBRUPTURE1',False);T.AddChampSup('LIBRUPTURE2',False);
                        T.AddChampSup('LIBRUPTURE3',False);T.AddChampSup('LIBRUPTURE4',False);
                        T.AddChampSup('HEURESR',False);T.AddChampSup('COUTSR',False);
                        T.AddChampSup('HEURESNR',False);T.AddChampSup('COUTSNR',False);
                        T.AddChampSup('HEURESNP',False);T.AddChampSup('COUTSNP',False);
                        T.PutValue('HEURESNR',0);
                        T.PutValue('COUTSNR',0);
                        If NbRupt > 3 then T.PutValue('RUPTURE4',TobRealise.Detail[i].GetValue(ChampRuptRea4));
                        If NbRupt > 2 then T.PutValue('RUPTURE3',TobRealise.Detail[i].GetValue(ChampRuptRea3));
                        If NbRupt > 1 then T.PutValue('RUPTURE2',TobRealise.Detail[i].GetValue(ChampRuptRea2));
                        T.PutValue('RUPTURE1',TobRealise.Detail[i].GetValue(ChampRuptRea1));
                        T.PutValue('HEURESR',0);T.PutValue('COUTSR',0);
                        T.PutValue('HEURESNP',NbHeuresR);T.PutValue('COUTSNP',TotalRea);
                end;
        end;
        TobRealise.Free;
        TobNbInscRea.Free;
        For i := 0 to TobEtat.detail.Count-1 do
        begin
                If NbRupt > 3 then
                begin
                        Libelle := RechDom(Tablette4,TobEtat.Detail[i].GetValue('RUPTURE4'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE4',Libelle);
                end;
                If NbRupt > 2 then
                begin
                        Libelle := RechDom(Tablette3,TobEtat.Detail[i].GetValue('RUPTURE3'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE3',Libelle);
                end;
                If NbRupt > 1 then
                begin
                        Libelle := RechDom(Tablette2,TobEtat.Detail[i].GetValue('RUPTURE2'),False);
                        If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE2',Libelle);
                end;
                Libelle := RechDom(Tablette1,TobEtat.Detail[i].GetValue('RUPTURE1'),False);
                If Libelle <> '' then TobEtat.Detail[i].PutValue('LIBRUPTURE1',Libelle);
        end;
        TobEtat.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;RUPTURE4');
        StPages := '';
       {$IFDEF EAGLCLIENT}
        StPages := AglGetCriteres (Pages, FALSE);
        {$ENDIF}
        If GetCheckBoxState('FLISTE') = CbChecked then LanceEtatTOB('E','PFO','PCC',TobEtat,True,True,False,Pages,' ','',False,0,StPages)
        else LanceEtatTOB('E','PFO','PCC',TobEtat,True,False,False,Pages,' ','',False,0,StPages);

        TobEtat.free;
end;

procedure TOF_PGFORCOMPCOLL.ChangeMillesime (Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+GetControlText('MILLESIME')+'"',True);
        If Not Q.eof then
        begin
                SetControlText('DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                SetControlText('DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

Function TOF_PGFORCOMPCOLL.ChampRuptureForm(TypeRupt,ValeurRupt : String) : String;
var Tablette : String;
    j : Integer;
begin
        If ValeurRupt = 'PSS_CODESTAGE' then Tablette := 'PGSTAGEFORM';
        If ValeurRupt = 'PFO_SALARIE' then Tablette := 'PGSALARIE';
        If ValeurRupt = 'PFO_ETABLISSEMENT' then Tablette := 'TTETABLISSEMENT';
        If ValeurRupt = 'PFO_LIBEMPLOIFOR' then Tablette := 'PGLIBEMPLOI';
        If ValeurRupt = 'PSS_NATUREFORM' then Tablette := 'PGNATUREFORM';
        If ValeurRupt = 'PSS_LIEUFORM' then Tablette := 'PGLIEUFORMATION';
        For j := 1 to 8 do
        begin
                If ValeurRupt = 'PSS_FORMATION'+IntToStr(j) then
                begin
                        Tablette := 'PGFORMATION'+IntToStr(j);
                end;
        end;
        For j := 1 to 4 do
        begin
                If ValeurRupt = 'PFO_TRAVAILN'+IntToStr(j) then
                begin
                Tablette := 'PGTRAVAILN'+IntToStr(j);
                end;
        end;
        Result := Tablette;
end;

{TOF_PGEDITFORCOMPDETAIL}
procedure TOF_PGEDITFORCOMPDETAIL.OnArgument (S : String ) ;
var Num,i : Integer;
    Millesime,Combo : THValComboBox;
    DD,DF : TDateTime;
begin
  Inherited ;
        SetControlCaption('LIBSTAGE','');
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PFI_FORMATION'+IntToStr(Num)),GetControl ('TPFI_FORMATION'+IntToStr(Num)));
        end;
        Millesime := THValComboBox(GetControl('PST_MILLESIME'));
        If Millesime <> Nil Then Millesime.Value := RendMillesimeRealise(DD,DF);
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('THVALRUPTURE'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                        Combo.Items.Add('<<Aucun>>');
                        Combo.Values.Add('');
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('PFI_CODESTAGE');
                        Combo.Items.Add('Responsable formation');
                        Combo.Values.Add('PFI_RESPONSFOR');
                        Combo.Items.Add('Etablissement');
                        Combo.Values.Add('PFI_ETABLISSEMENT');
                        Combo.Items.Add('Libellé emploi');
                        Combo.Values.Add('PFI_LIBEMPLOIFOR');
                        Combo.Items.Add('Type (Interne/Externe)');
                        Combo.Values.Add('PST_NATUREFORM');
                        Combo.Items.Add('Motif inscription');
                        Combo.Values.Add('PFI_MOTIFINSCFOR');
                        Combo.Items.Add('Etat inscription');
                        Combo.Values.Add('PFI_ETATINSCFOR');
                        Combo.Items.Add('Motif décision');
                        Combo.Values.Add('PFI_MOTIFETATINSC');
                        For Num  :=  1 to VH_Paie.NBFormationLibre do
                        begin
                                If Num=1 Then begin Combo.Items.Add(VH_Paie.FormationLibre1); Combo.Values.Add('PFI_FORMATION1');end;
                                If Num=2 Then begin Combo.Items.Add(VH_Paie.FormationLibre2); Combo.Values.Add('PFI_FORMATION2');end;
                                If Num=3 Then begin Combo.Items.Add(VH_Paie.FormationLibre3); Combo.Values.Add('PFI_FORMATION3');end;
                                If Num=4 Then begin Combo.Items.Add(VH_Paie.FormationLibre4); Combo.Values.Add('PFI_FORMATION4');end;
                                If Num=5 Then begin Combo.Items.Add(VH_Paie.FormationLibre5); Combo.Values.Add('PFI_FORMATION5');end;
                                If Num=6 Then begin Combo.Items.Add(VH_Paie.FormationLibre6); Combo.Values.Add('PFI_FORMATION6');end;
                                If Num=7 Then begin Combo.Items.Add(VH_Paie.FormationLibre7); Combo.Values.Add('PFI_FORMATION7');end;
                                If Num=8 Then begin Combo.Items.Add(VH_Paie.FormationLibre8); Combo.Values.Add('PFI_FORMATION8');end;
                        end;
                end;
        end;
        SetControlProperty('THVALRUPTURE1','Value','');
        SetControlProperty('THVALRUPTURE2','Value','');
        SetControlProperty('THVALRUPTURE3','Value','');
        SetControlProperty('THVALRUPTURE4','Value','');
end;

procedure TOF_PGEDITFORCOMPDETAIL.OnUpdate ;
var StWhere,StOrder,StSelect : String;
    Pages : TPageControl;
    Rupt1,Rupt2,Rupt3,Rupt4 : String;
begin
  Inherited ;
        Pages := TPageControl(GetControl('Pages'));
        StWhere := RecupWhereCritere(Pages);
        StSelect := 'SELECT PST_CODESTAGE,PST_FORMATION1 '+
        'FROM STAGE';
        Rupt1 := GetControlText('THVALRUPTURE1');
        Rupt2 := GetControlText('THVALRUPTURE2');
        Rupt3 := GetControlText('THVALRUPTURE3');
        Rupt4 := GetControlText('THVALRUPTURE4');
        SetControlText('XX_RUPTURE1',Rupt1);
        SetControlText('XX_RUPTURE2',Rupt2);
        SetControlText('XX_RUPTURE3',Rupt3);
        SetControlText('XX_RUPTURE4',Rupt4);
        StOrder := '';
        If Rupt1 <> '' then StOrder := 'ORDER BY '+Rupt1;
        If Rupt2 <> '' then StOrder := StOrder+','+Rupt2;
        If Rupt3 <> '' then StOrder := StOrder+','+Rupt3;
        If Rupt4 <> '' then StOrder := StOrder+','+Rupt4;
        TFQRS1(Ecran).WhereSQL := StSelect +' '+StWhere+' '+StOrder;
end;


Initialization
registerclasses([TOF_PGEDITFORGESTIND,TOF_PGEDITFORGESTPREV,TOF_PGFORSUIVIGESTCOLL,TOF_PGEDITFORSUIVIINV,TOF_PGFORCOMPCOLL,TOF_PGEDITFORCOMPDETAIL]);
end.
