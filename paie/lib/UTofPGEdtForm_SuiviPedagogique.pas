{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 01/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITCOMPARATIFFOR ()
Mots clefs ... : TOF;PGEDITCOMPARATIFFOR
Contient les sources TOF pour les éditions de la formation
*****************************************************************
PT1  | 24/04/2003 | V_42  | JL | Développement pour CWAS
PT2  | 09/03/2005 | V_60  | JL | Ajout etats libres pour CEGID
---  | 20/03/2006 |       | JL | modification clé annuaire ----
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT3  | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
     |            |       |    |  pour les critères code salarié uniquement
PT4  | 26/10/2007 | V_80  | FL | Emanager / Report / Adaptation du suivi animateurs
PT5  | 25/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT6  | 03/04/2008 | V_803 | FL | Adaptation pour partage formation
PT7  | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue + Gestion elipsis salarié et responsable
PT8  | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT9  | 04/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
}

unit UTofPGEdtForm_SuiviPedagogique;

interface                                                 
Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtREtat,
{$ELSE}
     eQRS1,UtilEAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,LookUp,
     PGOutilsFormation,P5Def,PGEditOutils,PGEditOutils2,HQry,HTB97,Utob,ParamSoc,PGOutils2,pgOutils ;

Type
  TOF_PGEDITSALARIESFORMATION = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;
    private
    Arg : String;
    procedure ChangeMillesime(Sender : TObject);
    procedure ChoixRuptFor(Sender : TObject);
    procedure ChoixruptSal(Sender : TObject);
    procedure CreerRuptSal(Num : Integer);
    procedure CreerRuptFor(Num : Integer);
    {$IFDEF EMANAGER}
    procedure SalarieElipsisClick(Sender : TObject);
    {$ENDIF}
  end ;

  //Suivi formation - Participants
  TOF_PGEDITCOLLECITVEFORMATION  = Class (TOF)
    procedure OnClose                  ; override ;
    procedure OnUpdate;              override;
    procedure OnArgument(S:String); override;
    private
    {$IFDEF EMANAGER}
    TypeUtilisat : String;
    {$ENDIF}
    TobEtat,T:Tob;
//    {$IFDEF EMANAGER} //PT6
    procedure SalarieElipsisClick(Sender : TObject);
//    {$ENDIF}
//    procedure BEtatClick(Sender : TObject);
    procedure ConstruireTob ();
    Function RendTablette(Champ : String) : String;
    procedure ChangeMillesime (Sender : TObject);
//    {$IFDEF EMANAGER} //PT6
    procedure RespElipsisClick(Sender : TObject);
//    {$ENDIF}
  end ;
  TOF_PGEDITANIMATEURS = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate ; override ; 
    Private
    procedure ChangeMillesime(Sender : TObject);
    procedure ChoixRupture(Sender : TObject);
    {$IFDEF EMANAGER}
    procedure SalarieElipsisClick(Sender : TObject);
    {$ENDIF}
  end;
  
  // Suivi animation interne
  TOF_PGEDITFORANIMATEUR = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    LeMenu{$IFDEF EMANAGER},TypeUtilisat{$ENDIF} : String; //PT4
    {$IFNDEF EMANAGER}
    procedure ExitEdit(Sender : TObject);
    {$ENDIF}
    procedure SalarieElipsisClick (Sender : TObject); //PT4
    //{$ENDIF}
  end ;
  
  // Suivi formation - Animateurs / stagiaires
  TOF_EDITSTAGIAIREANIMATEUR = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;
  private
    LeMenu : String;
    procedure SalarieElipsisClick(Sender : TObject); //PT6
  end;


implementation
Uses GalOutil;

{TOF_PGEDITSALARIESFORMATION}
// Edition individuel des salariés pour formation
procedure TOF_PGEDITSALARIESFORMATION.OnUpdate;
var OrderBy,Where,SQL : String;
   Pages : TPageControl;
   i : Integer;
   Combo : THValCombobox;
begin
Inherited ;
 {       For i := 1 To 8 do
        begin
                If GetCheckBoxState('CRUPT'+IntToStr(i))=CbChecked Then
                begin
                        SetControlText('XX_RUPTURE2','PST_FORMATION'+IntToStr(i));
                        OrderByFor := 'PST_FORMATION'+IntToStr(i);
                end;
        end;
        For i := 1 To 4 do
        begin
                If GetCheckBoxState('CRUPTSAL'+IntToStr(i))=CbChecked Then
                begin
                        SetControlText('XX_RUPTURE1','PFO_TRAVAILN'+IntToStr(i));
                        OrderBySal := 'PFO_TRAVAILN'+IntToStr(i);
                end;
        end;
        If GetCheckBoxState('CRUPTSAL5')=CbChecked Then
        begin
                SetControlText('XX_RUPTURE1','PFO_CODESTAT');
                OrderBySal := 'PFO_CODESTAT';
        end;
        If GetCheckBoxState('CRUPTSAL6')=CbChecked Then
        begin
                SetControlText('XX_RUPTURE1','PFO_ETABLISSEMENT');
                OrderBySal := 'PFO_ETABLISSEMENT';
        end;
        Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        If GetCheckBoxState('CALPHA')=CbChecked Then TriSal := 'PFO_NOMSALARIE,PFO_SALARIE'
        Else TriSal := 'PFO_SALARIE';
        //Debut PT1
        If Arg ='CWAS' then
        begin
                If Where <> '' Then Where := Where+' AND ';
                Where := Where+'AND PFO_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")';
        end;
        //Fin PT1
        If Where <> '' Then Where := ' '+Where;
        OrderBy := '';
        If OrderBySal <> '' Then OrderBy := ' ORDER BY '+OrderBySal;
        If OrderBy <> '' Then orderBy := OrderBy+','+Trisal
        Else OrderBy := ' ORDER BY '+TriSal;
        If OrderByFor <> '' Then OrderBy := OrderBy+','+OrderByFor;
        SQL := 'SELECT PFO_SALARIE,PFO_CODESTAGE,PFO_NOMSALARIE,PFO_ORDRE,PFO_PRENOM,PFO_NBREHEURE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_LIEUFORM,'+
        'PFO_ETABLISSEMENT,PFO_CODESTAT,PFO_TRAVAILN1,PFO_TRAVAILN2,PFO_TRAVAILN3,PFO_TRAVAILN4,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4'+
        ',PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8  '+
        'FROM FORMATIONS LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME';
        TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;         }
        Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        If Arg = 'CWAS' then
        begin
        If Where <> '' then Where := Where + ' AND PFO_RESPONSFOR="'+V_PGI.UserSalarie+'"'
        else Where := 'WHERE PFO_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        end;
        For i := 1 to 4 do
        begin
                Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
                If Combo.Value <> '' Then
                begin
                        If Combo.Value = 'F1' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION1');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION1';
                        end;
                        If Combo.Value = 'F2' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION2');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION2';
                        end;
                        If Combo.Value = 'F3' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION3');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION3';
                        end;
                        If Combo.Value = 'F4' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION4');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION4';
                        end;
                        If Combo.Value = 'F5' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION5');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION5';
                        end;
                        If Combo.Value = 'F6' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION6');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION6';
                        end;
                        If Combo.Value = 'F7' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION7');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION7';
                        end;
                        If Combo.Value = 'F8' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_FORMATION8');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_FORMATION8';
                        end;
                        If Combo.Value = 'T1' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_TRAVAILN1');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PSA_TRAVAILN1';
                        end;
                        If Combo.Value = 'T2' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_TRAVAILN2');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PSA_TRAVAILN2';
                        end;
                        If Combo.Value = 'T3' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_TRAVAILN3');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PSA_TRAVAILN3';
                        end;
                        If Combo.Value = 'T4' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_TRAVAILN4');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PSA_TRAVAILN4';
                        end;
                        If Combo.Value = 'CS' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_CODESTAT');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PSA_CODESTAT';
                        end;
                        If Combo.Value = 'NF' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PST_NATUREFORM');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PST_NATUREFORM';
                        end;
                        If Combo.Value = 'STA' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_CODESTAGE');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PFO_CODESTAGE';
                        end;
                        If Combo.Value = 'ETAB' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_ETABLISSEMENT');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PSA_ETABLISSEMENT';
                        end;
                        If Combo.Value = 'RESP' Then
                        begin
                                SetControlText('XX_RUPTURE'+IntToStr(i),'PFO_RESPONSFOR');
                                If OrderBy <> '' Then OrderBy := OrderBy+',';
                                OrderBy := OrderBy+'PFO_RESPONSFOR';
                        end;
                end
                Else SetControlText('XX_RUPTURE'+IntToStr(i),'');
        end;
        If OrderBy <> '' Then OrderBy := OrderBy+',PFO_SALARIE'
        Else OrderBy := 'PFO_SALARIE';
        OrderBy := 'ORDER BY '+OrderBy;
         SQL := 'SELECT PFO_RESPONSFOR,PFO_SALARIE,PFO_CODESTAGE,PFO_NOMSALARIE,PFO_ORDRE,PFO_PRENOM,PFO_NBREHEURE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_LIEUFORM,'+
        'PFO_ETABLISSEMENT,PFO_CODESTAT,PFO_TRAVAILN1,PFO_TRAVAILN2,PFO_TRAVAILN3,PFO_TRAVAILN4,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4'+
        ',PST_NATUREFORM,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM  '+
        'FROM FORMATIONS LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME ';
        
	    //PT8
	    // N'affiche pas les salariés confidentiels
	    If Where <> '' Then Where := Where + ' AND ';
	    Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
        
        TFQRS1(Ecran).WhereSQL := SQL+Where+' '+OrderBy;
end;

procedure TOF_PGEDITSALARIESFORMATION.OnArgument (S : String ) ;
var Num,i : Integer;
    Min,Max : String;
    Combo : THValComboBox;
    Check : TCheckBox;
    Q : TQuery;
    Edit : THEdit;
begin
  Inherited ;
        Edit := THEdit(GetControl('PFO_SALARIE'));
        If Edit <> nil then
        begin
                Edit.Operateur := Superieur;
                {$IFDEF EMANAGER}
                Edit.OnElipsisClick := SalarieElipsisClick;
                {$ENDIF}
        end;
        Edit := THEdit(GetControl('PFO_SALARIE_'));
        begin
                Edit.Operateur := Inferieur;
                {$IFDEF EMANAGER}
                Edit.OnElipsisClick := SalarieElipsisClick;
                {$ENDIF}
        end;
        Arg := ReadTokenPipe(S,';');
        For Num := 1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        For Num := 1 To 8 do
        begin
                Combo := THValComboBox(GetControl('PST_FORMATION'+IntToStr(Num)));
                If Combo.Visible=True Then SetControlVisible('CRUPT'+IntToStr(Num),True);
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
                If GetControlVisible('PFO_TRAVAILN'+IntToStr(Num))=True Then SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
        end;
        VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
        For Num := 1 to 6 do
        begin
                Check := TCheckBox(GetControl('CRUPTSAL'+IntToStr(Num)));
                If Check <> Nil Then Check.OnClick := ChoixRuptSal;
        end;
        For Num := 1 to 8 do
        begin
                Check := TCheckBox(GetControl('CRUPT'+IntToStr(Num)));
                If Check <> Nil Then Check.OnClick := ChoixRuptFor;
        end;
        If GetControlVisible('PFO_CODESTAT')=True Then SetControlVisible('CRUPTSAL5',True);
        RecupMinMaxTablette('PG','EXERFORMATION','PFE_MILLESIME',Min,Max);
        SetControlProperty('MILLESIME','Value',Min);
        SetControlProperty('MILLESIME_','Value',Max);
        Q := OpenSQL('SELECT PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_MILLESIME="'+Min+'"',true);
        If Not Q.eof Then SetControlText('PFO_DATEDEBUT',Q.FindField('PFE_DATEDEBUT').AsString);
        Ferme(Q);
        Q := OpenSQL('SELECT PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+Max+'"',true);
        If Not Q.eof Then SetControlText('PFO_DATEFIN',Q.FindField('PFE_DATEFIN').AsString);
        Ferme(Q);
        Combo := THValComboBox(GetControl('MILLESIME'));
        If Combo <> Nil then Combo.OnChange := ChangeMillesime;
        Combo := THValComboBox(GetControl('MILLESIME_'));
        If Combo <> Nil then Combo.OnChange := ChangeMillesime;
        SetControltext('XX_RUPTURE1','');
        SetControltext('XX_RUPTURE2','');
        SetControlCaption('LIBSTAGE','');
        For i := 1 to 4 do
        begin
                Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
                Combo.Items.add ('<<AUCUN>>');
                Combo.Values.add('');
                Combo.Items.add ('Formations');
                Combo.Values.add('STA');
                Combo.Items.add ('Etablissements');
                Combo.Values.add('ETAB');
                Combo.Items.add ('Responsables');
                Combo.Values.add('RESP');
                For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                begin
                        if Num = 1 then Combo.Values.add('F1')
                        else if Num = 2 then Combo.Values.add ('F2')
                        else if Num = 3 then Combo.Values.add ('F3')
                        else if Num = 4 then Combo.Values.add ('F4')
                        else if Num = 5 then Combo.Values.add ('F5')
                        else if Num = 6 then Combo.Values.add ('F6')
                        else if Num = 7 then Combo.Values.add ('F7')
                        else Combo.Values.add ('F8');
                        if Num = 1 then Combo.Items.add (VH_Paie.FormationLibre1)
                        else if Num = 2 then Combo.Items.add (VH_Paie.FormationLibre2)
                        else if Num = 3 then Combo.Items.add (VH_Paie.FormationLibre3)
                        else if Num = 4 then Combo.Items.add (VH_Paie.FormationLibre4)
                        else if Num = 5 then Combo.Items.add (VH_Paie.FormationLibre5)
                        else if Num = 6 then Combo.Items.add (VH_Paie.FormationLibre6)
                        else if Num = 7 then Combo.Items.add (VH_Paie.FormationLibre7)
                        else Combo.Items.add  (VH_Paie.FormationLibre8);
                end;
                For Num  :=  1 to VH_Paie.PGNbreStatOrg do
                begin
                        if Num = 1 then Combo.Values.add ('T1')
                        else if Num = 2 then Combo.Values.add ('T2')
                        else if Num = 3 then Combo.Values.add ('T3')
                        else Combo.Values.add ('T4');
                        if Num = 1 then Combo.Items.add (VH_Paie.PGLibelleOrgStat1)
                        else if Num = 2 then Combo.Items.add (VH_Paie.PGLibelleOrgStat2)
                        else if Num = 3 then Combo.Items.add (VH_Paie.PGLibelleOrgStat3)
                        else Combo.Items.add (VH_Paie.PGLibelleOrgStat4);
                end;
                if VH_Paie.PGLibCodeStat  <>  '' then
                begin
                        Combo.Values.add ('CS');
                        Combo.Items.add (VH_Paie.PGLibCodeStat);
                end;
                Combo.Values.add('NF');
                Combo.Items.add('Nature de la formation');
        end;
end;


procedure TOF_PGEDITSALARIESFORMATION.ChangeMillesime(Sender : TObject);
var Q : TQuery;
    Combo : THValComboBox;
begin
        Combo := THValComboBox(GetControl('MILLESIME'));
        Q := OpenSQL('SELECT PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_MILLESIME="'+Combo.Value+'"',true);
        If Not Q.eof Then SetControlText('PFO_DATEDEBUT',Q.FindField('PFE_DATEDEBUT').AsString);
        Ferme(Q);
        Combo := THValComboBox(GetControl('MILLESIME_'));
        Q := OpenSQL('SELECT PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+Combo.Value+'"',true);
        If Not Q.eof Then SetControlText('PFO_DATEFIN',Q.FindField('PFE_DATEFIN').AsString);
        Ferme(Q);
end;

procedure TOF_PGEDITSALARIESFORMATION.ChoixRuptFor(Sender : TObject);
begin
        If TCheckBox(Sender).Name = 'CRUPT1' Then CreerRuptFor(1);
        If TCheckBox(Sender).Name = 'CRUPT2' Then CreerRuptFor(2);
        If TCheckBox(Sender).Name = 'CRUPT3' Then CreerRuptFor(3);
        If TCheckBox(Sender).Name = 'CRUPT4' Then CreerRuptFor(4);
        If TCheckBox(Sender).Name = 'CRUPT5' Then CreerRuptFor(5);
        If TCheckBox(Sender).Name = 'CRUPT6' Then CreerRuptFor(6);
        If TCheckBox(Sender).Name = 'CRUPT7' Then CreerRuptFor(7);
        If TCheckBox(Sender).Name = 'CRUPT8' Then CreerRuptFor(8);
end;

procedure TOF_PGEDITSALARIESFORMATION.CreerRuptFor(Num : Integer);
var i : Integer;
begin
        If GetCheckBoxState('CRUPT'+IntToStr(Num))=CbChecked Then
        begin
                SetControlText('XX_RUPTURE2','PST_FORMATION'+IntToStr(Num));
                For i := 1 to 7 do
                begin
                        If i <> Num Then SetControlChecked('CRUPT'+IntToStr(i),false);
                end;
        end
        Else SetControlText('XX_RUPTURE2','');
end;

procedure TOF_PGEDITSALARIESFORMATION.ChoixRuptSal(Sender : TObject);
begin
        If TCheckBox(Sender).Name = 'CRUPTSAL1' Then CreerRuptSal(1);
        If TCheckBox(Sender).Name = 'CRUPTSAL2' Then CreerRuptSal(2);
        If TCheckBox(Sender).Name = 'CRUPTSAL3' Then CreerRuptSal(3);
        If TCheckBox(Sender).Name = 'CRUPTSAL4' Then CreerRuptSal(4);
        If TCheckBox(Sender).Name = 'CRUPTSAL5' Then CreerRuptSal(5);
        If TCheckBox(Sender).Name = 'CRUPTSAL6' Then CreerRuptSal(6);
end;

procedure TOF_PGEDITSALARIESFORMATION.CreerRuptSal(Num : Integer);
var i : Integer;
    Champ,ChampCB : String;
begin
        ChampCB := 'CRUPTSAL'+IntToStr(Num);
        If Num<5  Then Champ := 'PFO_TRAVAILN'+IntToStr(Num)
        Else
        begin
                If Num=5 Then Champ := 'PFO_CODESTAT';
                If Num=6 Then Champ := 'PFO_ETABLISSEMENT';
        end;
        If GetCheckBoxState(ChampCB)=CbChecked Then
        begin
                SetControlText('XX_RUPTURE1',Champ);
                For i := 1 to 6 do
                begin
                        If i <> Num Then SetControlChecked('CRUPTSAL'+IntToStr(i),false);
                end;
        end
        Else SetControlText('XX_RUPTURE1','');
end;

{$IFDEF EMANAGER}
procedure TOF_PGEDITSALARIESFORMATION.SalarieElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        StOrder := 'PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
        LookupList(THEdit(Sender),'Liste des stages','SALARIES','PSA_SALARIE LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;
{$ENDIF}

{TOF_PGEDITCOLLECITVEFORMATION}
procedure TOF_PGEDITCOLLECITVEFORMATION.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
     If GetControlText('MILLESIME')= '' then
     begin
          SetControlText('PSS_DATEDEBUT',DateToStr(IDate1900));
          SetControlText('PSS_DATEFIN',DateToStr(IDate1900));
     end
     else
     begin
          Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+GetControlText('MILLESIME')+'"',True);
          If Not Q.Eof then
          begin
               SetControlText('PSS_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
               SetControlText('PSS_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
          end;
          Ferme(Q);
     end;
end;

procedure TOF_PGEDITCOLLECITVEFORMATION.OnUpdate;
var  TobAnim,TobSta : Tob;
     Rupt1,Rupt2,Rupt3,Rupt4,Where,WhereAnim,ListeChamp,StPages : String;
     Pages : TPageControl;
     i : Integer;
     Q : TQuery;
     Tablette1,tablette2,tablette3,tablette4 : String;
     {$IFDEF EMANAGER}
     MultiNiveau : Boolean;
     {$ENDIF}
     Req : String; //PT6
begin
     Inherited ;
     	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS')); //PT6
     	
        Pages := TPageControl(GetControl('Pages'));
        Rupt1 := '';
        Rupt2 := '';
        Rupt3 := '';
        Rupt4 := '';
        Where := RecupWhereCritere(Pages);
        If Where <> '' then WhereAnim := Where + ' AND PAN_SALARIE<>""'
        else WhereAnim := 'WHERE PAN_SALARIE <> ""';
        If GetControlText('VRUPT4') <> '' then
        begin
                Rupt4 := GetControltext('VRUPT4');
                Rupt3 := GetControltext('VRUPT3');
                Rupt2 := GetControltext('VRUPT2');
                Rupt1 := GetControltext('VRUPT1');
                SetControlText('XX_RUPTURE1','RUPTURE1');
                SetControlText('XX_RUPTURE2','RUPTURE2');
                SetControlText('XX_RUPTURE3','RUPTURE3');
                SetControlText('XX_RUPTURE4','RUPTURE4');
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
                        end
                        Else
                        begin
                                Rupt1 := GetControltext('VRUPT1');
                                SetControlText('XX_RUPTURE1','RUPTURE1');
                                SetControlText('XX_RUPTURE2','');
                                SetControlText('XX_RUPTURE3','');
                                SetControlText('XX_RUPTURE4','');
                        end;
                end;
        end;
        tablette1 := '';tablette2 := '';tablette3 := '';tablette4 := '';
        Tablette1 := RendTablette(Rupt1);
        Tablette2 := RendTablette(Rupt2);
        Tablette3 := RendTablette(Rupt3);
        Tablette4 := RendTablette(Rupt4);
        If Rupt1 <> '' then ListeChamp := ','+Rupt1;
        If Rupt2 <> '' then ListeChamp := ListeChamp+','+Rupt2;
        If Rupt3 <> '' then ListeChamp := ListeChamp+','+Rupt3;
        If Rupt4 <> '' then ListeChamp := ListeChamp+','+Rupt4;
        If TobEtat <> Nil then FreeAndNil(TobEtat);
        TobEtat := Tob.Create('Les salariés pour édition',Nil,-1);
        
        {$IFDEF EMANAGER}
        MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
        If TypeUtilisat = 'R' then
        begin
             If MultiNiveau then WhereAnim := WhereAnim + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
             ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
             else WhereAnim := WhereAnim + ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        end
        else WhereAnim := WhereAnim + ' AND PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
        {$ENDIF}
             
        
	    //PT8
	    // N'affiche pas les salariés confidentiels
	    If WhereAnim <> '' Then WhereAnim := WhereAnim + ' AND ';
	    WhereAnim := WhereAnim + '(PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';  //PT9
			    
        Req := 'SELECT PSA_SALARIE,PST_LIBELLE,PAN_NBREHEURE,PAN_DATEDEBUT,PAN_DATEFIN'+ListeChamp+' FROM SESSIONANIMAT '+
        'LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE '+
        'LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME '+
        'LEFT JOIN STAGE ON PAN_CODESTAGE=PST_CODESTAGE AND PAN_MILLESIME=PST_MILLESIME '+
        'LEFT JOIN DEPORTSAL ON PAN_SALARIE=PSE_SALARIE '+WhereAnim;

		//PT6 - Début
        If PGBundleHierarchie Then
        Begin
        	Req := StringReplace(Req, 'PSA_SALARIE', 'PSI_INTERIMAIRE', [rfReplaceAll,rfIgnoreCase]);
        	Req := StringReplace(Req, 'PSA_', 'PSI_', [rfReplaceAll,rfIgnoreCase]);
        	Req := StringReplace(Req, 'SALARIES',    'INTERIMAIRES',    [rfReplaceAll,rfIgnoreCase]);
        End;
        //PT6 - Fin
        
        Q := OpenSQL(Req,True);
        TobAnim := Tob.Create('Les animateurs',Nil,-1);
        TobAnim.LoadDetailDB('Les animateurs','','',Q,False);
        Ferme(Q);
        
        For i := 0 to TobAnim.Detail.Count -1 do
        begin
                ConstruireTob;
                If Rupt1 <> '' then
                begin
                        T.PutValue('RUPTURE1',TobAnim.Detail[i].GetValue(Rupt1));
                        T.PutValue('LIBRUPTURE1',RechDom(Tablette1,TobAnim.Detail[i].GetValue(Rupt1),False));
                end
                else
                begin
                        T.PutValue('RUPTURE1','');
                        T.PutValue('LIBRUPTURE1','');
                end;
                If Rupt2 <> '' then
                begin
                        T.PutValue('RUPTURE2',TobAnim.Detail[i].GetValue(Rupt2));
                        T.PutValue('LIBRUPTURE2',RechDom(Tablette2,TobAnim.Detail[i].GetValue(Rupt2),False));
                end
                else
                begin
                        T.PutValue('RUPTURE2','');
                        T.PutValue('LIBRUPTURE2','');
                end;
                If Rupt3 <> '' then
                begin
                        T.PutValue('RUPTURE3',TobAnim.Detail[i].GetValue(Rupt3));
                        T.PutValue('LIBRUPTURE3',RechDom(Tablette3,TobAnim.Detail[i].GetValue(Rupt3),False));
                end
                else
                begin
                        T.PutValue('RUPTURE3','');
                        T.PutValue('LIBRUPTURE3','');
                end;
                If Rupt4 <> '' then
                begin
                        T.PutValue('RUPTURE4',TobAnim.Detail[i].GetValue(Rupt4));
                        T.PutValue('LIBRUPTURE4',RechDom(Tablette4,TobAnim.Detail[i].GetValue(Rupt4),False));
                end
                else
                begin
                        T.PutValue('RUPTURE4','');
                        T.PutValue('LIBRUPTURE4','');
                end;
                T.PutValue('TYPE','Animateurs');
                T.PutValue('LIBSTAGE',TobAnim.Detail[i].GetValue('PST_LIBELLE'));
                T.PutValue('DATEDEBUT',TobAnim.Detail[i].GetValue('PAN_DATEDEBUT'));
                T.PutValue('DATEFIN',TobAnim.Detail[i].GetValue('PAN_DATEFIN'));
                T.PutValue('NBREHEURE',TobAnim.Detail[i].GetValue('PAN_NBREHEURE'));
                //PT6 - Début
                If Not PGBundleHierarchie Then
                	T.PutValue('PSA_SALARIE',TobAnim.Detail[i].GetValue('PSA_SALARIE'))
                Else
                	T.PutValue('PSA_SALARIE',TobAnim.Detail[i].GetValue('PSI_INTERIMAIRE'));
                //PT6 - Fin
        end;
        TobAnim.Free;
        
        {$IFDEF EMANAGER}
        If TypeUtilisat = 'R' then
        begin
             If MultiNiveau then
             begin
                  If Where <> '' then Where := Where + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                  ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
                  else Where := ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                  ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
             end
             else
             begin
                  If Where <> '' then Where := Where + ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'
                  else Where := ' WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
             end;
        end
        else
        begin
             If Where <> '' then Where := Where + ' AND PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))'
             else Where := ' WHERE PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';

        end;
        {$ENDIF}
        
	    //PT8
	    // N'affiche pas les salariés confidentiels
	    If Where <> '' Then Where := Where + ' AND ';
	    Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
        
        Req := 'SELECT PSA_SALARIE,PST_LIBELLE,PFO_NBREHEURE,PFO_DATEDEBUT,PFO_DATEFIN'+ListeChamp+' FROM FORMATIONS '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME '+
        'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME '+
        'LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE '+Where;
        
        //PT6 - Début
        If PGBundleHierarchie Then
        Begin
        	Req := StringReplace(Req, 'PSA_SALARIE', 'PSI_INTERIMAIRE', [rfReplaceAll,rfIgnoreCase]);
        	Req := StringReplace(Req, 'PSA_', 'PSI_', [rfReplaceAll,rfIgnoreCase]);
        	Req := StringReplace(Req, 'SALARIES',    'INTERIMAIRES',    [rfReplaceAll,rfIgnoreCase]);
        End;
        //PT6 - Fin
        
        Q := OpenSQL(Req,True);
        
        TobSta := Tob.Create('Les animateurs',Nil,-1);
        TobSta.LoadDetailDB('Les animateurs','','',Q,False);
        Ferme(Q);
        For i := 0 to TobSta.Detail.Count -1 do
        begin
                construireTob;
                If Rupt1 <> '' then
                begin
                        T.PutValue('RUPTURE1',TobSta.Detail[i].GetValue(Rupt1));
                        T.PutValue('LIBRUPTURE1',RechDom(Tablette1,TobSta.Detail[i].GetValue(Rupt1),False));
                end
                else
                begin
                        T.PutValue('RUPTURE1','');
                        T.PutValue('LIBRUPTURE1','');
                end;
                If Rupt2 <> '' then
                begin
                        T.PutValue('RUPTURE2',TobSta.Detail[i].GetValue(Rupt2));
                        T.PutValue('LIBRUPTURE2',RechDom(Tablette2,TobSta.Detail[i].GetValue(Rupt2),False));
                end
                else
                begin
                        T.PutValue('RUPTURE2','');
                        T.PutValue('LIBRUPTURE2','');
                end;
                If Rupt3 <> '' then
                begin
                        T.PutValue('RUPTURE3',TobSta.Detail[i].GetValue(Rupt3));
                        T.PutValue('LIBRUPTURE3',RechDom(Tablette3,TobSta.Detail[i].GetValue(Rupt3),False));
                end
                else
                begin
                        T.PutValue('RUPTURE3','');
                        T.PutValue('LIBRUPTURE3','');
                end;
                If Rupt4 <> '' then
                begin
                        T.PutValue('RUPTURE4',TobSta.Detail[i].GetValue(Rupt4));
                        T.PutValue('LIBRUPTURE4',RechDom(Tablette4,TobSta.Detail[i].GetValue(Rupt4),False));
                end
                else
                begin
                        T.PutValue('RUPTURE4','');
                        T.PutValue('LIBRUPTURE4','');
                end;
                T.PutValue('TYPE','Stagiaires');
                T.PutValue('LIBSTAGE',TobSta.Detail[i].GetValue('PST_LIBELLE'));
                T.PutValue('DATEDEBUT',TobSta.Detail[i].GetValue('PFO_DATEDEBUT'));
                T.PutValue('DATEFIN',TobSta.Detail[i].GetValue('PFO_DATEFIN'));
                T.PutValue('NBREHEURE',TobSta.Detail[i].GetValue('PFO_NBREHEURE'));
                //PT6 - Début
                If Not PGBundleHierarchie Then
                	T.PutValue('PSA_SALARIE',TobSta.Detail[i].GetValue('PSA_SALARIE'))
                Else
                	T.PutValue('PSA_SALARIE',TobSta.Detail[i].GetValue('PSI_INTERIMAIRE'));
                //PT6 - Fin
        end;
        TobSta.Free;
        TobEtat.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;RUPTURE4');
        StPages := '';
        {$IFDEF EAGLCLIENT}
        StPages := AglGetCriteres (Pages, FALSE);
        {$ENDIF}
//        If GetCheckBoxState('FLISTE') = CbChecked then LanceEtatTOB('E','PFO','PCF',TobEtat,True,True,False,Pages,' ','',False,0,StPages)
//        else LanceEtatTOB('E','PFO','PCF',TobEtat,True,False,False,Pages,' ','',False,0,StPages);
//        TobEtat.Free;
          TFQRS1(Ecran).CodeEtat:= 'PCF';
        TFQRS1(Ecran).LaTob:= TobEtat;
end;

procedure TOF_PGEDITCOLLECITVEFORMATION.OnClose;
begin
     If Assigned(TobEtat) Then FreeAndNil (TobEtat);
end;

procedure TOF_PGEDITCOLLECITVEFORMATION.OnArgument(S : String);
var Num,i : Integer;
    DD,DF : TDateTime;
    Millesime : String;
    //{$IFDEF EMANAGER} //PT6
    Edit : THEdit;
    //{$ENDIF}
    Combo : THValComboBox;
    Prefixe : String; //PT6
begin
  Inherited ;
        SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
        Millesime := RendMillesimeRealise(DD,DF);
        SetControltext('PSS_DATEDEBUT',DateToStr(DD));
        SetControltext('PSS_DATEFIN',DateToStr(DF));
        SetControlText('MILLESIME',Millesime);
        Combo := THValComboBox(GetControl('MILLESIME'));

        //PT6
        If PGBundleHierarchie Then
            Prefixe := 'PSI'
        Else
            Prefixe := 'PSA';

        If Combo <> Nil then Combo.OnChange := ChangeMillesime;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PSS_FORMATION'+IntToStr(Num)),GetControl ('TPSS_FORMATION'+IntToStr(Num)));
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                        If i>1 then
                        begin
                                Combo.Items.Add('<<Aucune>>');
                                Combo.Values.Add('');
                        end;
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('PSS_CODESTAGE');
                        Combo.Items.Add('Salarié');
                        //PT6
                        If PGBundleHierarchie Then
                        Combo.Values.Add('PSI_INTERIMAIRE')
                        Else
                        Combo.Values.Add('PSA_SALARIE');
                        Combo.Items.Add('Etablissement');
                        Combo.Values.Add(Prefixe+'_ETABLISSEMENT');
                        Combo.Items.Add('Libellé emploi');
                        Combo.Values.Add(Prefixe+'_LIBELLEEMPLOI');
                        Combo.Items.Add('Nature de la formation');
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
        end ;
        //PT6
        If PGBundleHierarchie Then
            SetControlText('VRUPT1','PSI_INTERIMAIRE')
        Else
            SetControlText('VRUPT1','PSA_SALARIE');
        SetControlText('VRUPT2','');
        SetControlText('VRUPT3','');
        SetControlText('VRUPT4','');
        SetControlCaption('LIBSTAGE','');
        SetControlCaption('LIBRESP','');
        {$IFDEF EMANAGER}
        SetControlChecked('CSAUT1',True);
        SetControlVisible('CMULTINIVEAU',True);
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        SetControlProperty('PRUPTURES','TabVisible',False);
        //SetControlVisible('TPSE_RESPONSFOR',False);
        //SetControlVisible('PSE_RESPONSFOR',False);
        Ecran.Caption := 'Edition du réalisé';
        UpdateCaption(Ecran);
        {$ELSE}
        If VH_Paie.PGRESPONSABLES = False then
        begin
                SetControlVisible('TPSE_RESPONSFOR',False);
                SetControlVisible('PSE_RESPONSFOR',False);
        end;
        {$ENDIF}
        //PT6 - Déplacement en dehors du IFDEF
        {$IFNDEF EMANAGER}
        If PGBundleHierarchie Then //PT7
        Begin
        {$ENDIF}
	        Edit := THEdit(GetControl('PSE_RESPONSFOR'));
	        If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick;
	        Edit := THEdit(GetControl('PSA_SALARIE'));
	        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	        Edit := THEdit(GetControl('PSA_SALARIE_'));
	        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	    {$IFNDEF EMANAGER}
	    End;
	    {$ENDIF}
        
    //PT6 - Début
    If PGBundleCatalogue Then //PT7
    Begin
    	If not PGDroitMultiForm then
			SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
       		SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PSS_PREDEFINI',False);
			SetControlText   ('PSS_PREDEFINI','');
			//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT6 //PT7
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT6
       		//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6 //PT7
       	End;
	end
	else
	begin
		SetControlVisible('PSS_PREDEFINI', False);
		SetControlVisible('TPSS_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT6 - Fin
end;

procedure TOF_PGEDITCOLLECITVEFORMATION.SalarieElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StWhere,StOrder : String;
{$ENDIF}
begin
	{$IFDEF EMANAGER}
     If typeUtilisat = 'R' then StWhere := 'PSE_CODESERVICE IN '+
     '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
     ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
     ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))'
     else StWhere := 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     If GetControlText('RESPONSFOR') <> '' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'"';
     If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     StOrder := 'PSA_SALARIE';
     StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
     LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
     {$ELSE}
     //PT6 - Début
    //If PGBundleInscFormation Then //PT7
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
     {$ENDIF}
end;

procedure TOF_PGEDITCOLLECITVEFORMATION.RespElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StOrder,StWhere : String;
    St : String;
{$ENDIF}
begin
        {$IFNDEF EMANAGER}
        //PT6
        ElipsisResponsableMultidos (Sender);
        //St := ' SELECT PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM FROM SALARIES ';
        //StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES)';
        //LookupList(THEdit(Sender),'Liste des responsables de formation','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,'', True,-1);
        {$ELSE}
        If TypeUtilisat = 'R' then StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
        else StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
        StOrder := 'PSI_LIBELLE';
        LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
        {$ENDIF}
end;



{procedure TOF_PGEDITCOLLECITVEFORMATION.BEtatClick(Sender : TObject);
var  TobAnim,TobSta : Tob;
     Rupt1,Rupt2,Rupt3,Rupt4,Where,WhereAnim,ListeChamp,StPages : String;
     Pages : TPageControl;
     i : Integer;
     Q : TQuery;
     Tablette1,tablette2,tablette3,tablette4 : String;
begin
        Pages := TPageControl(GetControl('Pages'));
        Rupt1 := '';
        Rupt2 := '';
        Rupt3 := '';
        Rupt4 := '';
        Where := RecupWhereCritere(Pages);
        If Where <> '' then WhereAnim := Where + ' AND PAN_SALARIE<>""'
        else WhereAnim := 'WHERE PAN_SALARIE <> ""';
        If GetControlText('VRUPT4') <> '' then
        begin
                Rupt4 := GetControltext('VRUPT4');
                Rupt3 := GetControltext('VRUPT3');
                Rupt2 := GetControltext('VRUPT2');
                Rupt1 := GetControltext('VRUPT1');
                SetControlText('XX_RUPTURE1','RUPTURE1');
                SetControlText('XX_RUPTURE2','RUPTURE2');
                SetControlText('XX_RUPTURE3','RUPTURE3');
                SetControlText('XX_RUPTURE4','RUPTURE4');
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
                        end
                        Else
                        begin
                                Rupt1 := GetControltext('VRUPT1');
                                SetControlText('XX_RUPTURE1','RUPTURE1');
                                SetControlText('XX_RUPTURE2','');
                                SetControlText('XX_RUPTURE3','');
                                SetControlText('XX_RUPTURE4','');
                        end;
                end;
        end;
        tablette1 := '';tablette2 := '';tablette3 := '';tablette4 := '';
        Tablette1 := RendTablette(Rupt1);
        Tablette2 := RendTablette(Rupt2);
        Tablette3 := RendTablette(Rupt3);
        Tablette4 := RendTablette(Rupt4);
        If Rupt1 <> '' then ListeChamp := ','+Rupt1;
        If Rupt2 <> '' then ListeChamp := ListeChamp+','+Rupt2;
        If Rupt3 <> '' then ListeChamp := ListeChamp+','+Rupt3;
        If Rupt4 <> '' then ListeChamp := ListeChamp+','+Rupt4;
        TobEtat := Tob.Create('Les salariés pour édition',Nil,-1);
        Q := OpenSQL('SELECT PSA_SALARIE,PST_LIBELLE,PAN_NBREHEURE,PAN_DATEDEBUT,PAN_DATEFIN'+ListeChamp+' FROM SESSIONANIMAT '+
        'LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE '+
        'LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME '+
        'LEFT JOIN STAGE ON PAN_CODESTAGE=PST_CODESTAGE AND PAN_MILLESIME=PST_MILLESIME '+
        'LEFT JOIN DEPORTSAL ON PAN_SALARIE=PSE_SALARIE '+WhereAnim,True);
        TobAnim := Tob.Create('Les animateurs',Nil,-1);
        TobAnim.LoadDetailDB('Les animateurs','','',Q,False);
        Ferme(Q);
        For i := 0 to TobAnim.Detail.Count -1 do
        begin
                ConstruireTob;
                If Rupt1 <> '' then
                begin
                        T.PutValue('RUPTURE1',TobAnim.Detail[i].GetValue(Rupt1));
                        T.PutValue('LIBRUPTURE1',RechDom(Tablette1,TobAnim.Detail[i].GetValue(Rupt1),False));
                end
                else
                begin
                        T.PutValue('RUPTURE1','');
                        T.PutValue('LIBRUPTURE1','');
                end;
                If Rupt2 <> '' then
                begin
                        T.PutValue('RUPTURE2',TobAnim.Detail[i].GetValue(Rupt2));
                        T.PutValue('LIBRUPTURE2',RechDom(Tablette2,TobAnim.Detail[i].GetValue(Rupt2),False));
                end
                else
                begin
                        T.PutValue('RUPTURE2','');
                        T.PutValue('LIBRUPTURE2','');
                end;
                If Rupt3 <> '' then
                begin
                        T.PutValue('RUPTURE3',TobAnim.Detail[i].GetValue(Rupt3));
                        T.PutValue('LIBRUPTURE3',RechDom(Tablette3,TobAnim.Detail[i].GetValue(Rupt3),False));
                end
                else
                begin
                        T.PutValue('RUPTURE3','');
                        T.PutValue('LIBRUPTURE3','');
                end;
                If Rupt4 <> '' then
                begin
                        T.PutValue('RUPTURE4',TobAnim.Detail[i].GetValue(Rupt4));
                        T.PutValue('LIBRUPTURE4',RechDom(Tablette4,TobAnim.Detail[i].GetValue(Rupt4),False));
                end
                else
                begin
                        T.PutValue('RUPTURE4','');
                        T.PutValue('LIBRUPTURE4','');
                end;
                T.PutValue('TYPE','Animateurs');
                T.PutValue('LIBSTAGE',TobAnim.Detail[i].GetValue('PST_LIBELLE'));
                T.PutValue('DATEDEBUT',TobAnim.Detail[i].GetValue('PAN_DATEDEBUT'));
                T.PutValue('DATEFIN',TobAnim.Detail[i].GetValue('PAN_DATEFIN'));
                T.PutValue('NBREHEURE',TobAnim.Detail[i].GetValue('PAN_NBREHEURE'));
                T.PutValue('PSA_SALARIE',TobAnim.Detail[i].GetValue('PSA_SALARIE'));
        end;
        TobAnim.Free;
        Q := OpenSQL('SELECT PSA_SALARIE,PST_LIBELLE,PFO_NBREHEURE,PFO_DATEDEBUT,PFO_DATEFIN'+ListeChamp+' FROM FORMATIONS '+
        'LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE '+
        'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME '+
        'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME '+
        'LEFT JOIN DEPORTSAL ON PFO_SALARIE=PSE_SALARIE '+Where,True);
        TobSta := Tob.Create('Les animateurs',Nil,-1);
        TobSta.LoadDetailDB('Les animateurs','','',Q,False);
        Ferme(Q);
        For i := 0 to TobSta.Detail.Count -1 do
        begin
                construireTob;
                If Rupt1 <> '' then
                begin
                        T.PutValue('RUPTURE1',TobSta.Detail[i].GetValue(Rupt1));
                        T.PutValue('LIBRUPTURE1',RechDom(Tablette1,TobSta.Detail[i].GetValue(Rupt1),False));
                end
                else
                begin
                        T.PutValue('RUPTURE1','');
                        T.PutValue('LIBRUPTURE1','');
                end;
                If Rupt2 <> '' then
                begin
                        T.PutValue('RUPTURE2',TobSta.Detail[i].GetValue(Rupt2));
                        T.PutValue('LIBRUPTURE2',RechDom(Tablette2,TobSta.Detail[i].GetValue(Rupt2),False));
                end
                else
                begin
                        T.PutValue('RUPTURE2','');
                        T.PutValue('LIBRUPTURE2','');
                end;
                If Rupt3 <> '' then
                begin
                        T.PutValue('RUPTURE3',TobSta.Detail[i].GetValue(Rupt3));
                        T.PutValue('LIBRUPTURE3',RechDom(Tablette3,TobSta.Detail[i].GetValue(Rupt3),False));
                end
                else
                begin
                        T.PutValue('RUPTURE3','');
                        T.PutValue('LIBRUPTURE3','');
                end;
                If Rupt4 <> '' then
                begin
                        T.PutValue('RUPTURE4',TobSta.Detail[i].GetValue(Rupt4));
                        T.PutValue('LIBRUPTURE4',RechDom(Tablette4,TobSta.Detail[i].GetValue(Rupt4),False));
                end
                else
                begin
                        T.PutValue('RUPTURE4','');
                        T.PutValue('LIBRUPTURE4','');
                end;
                T.PutValue('TYPE','Stagiaires');
                T.PutValue('LIBSTAGE',TobSta.Detail[i].GetValue('PST_LIBELLE'));
                T.PutValue('DATEDEBUT',TobSta.Detail[i].GetValue('PFO_DATEDEBUT'));
                T.PutValue('DATEFIN',TobSta.Detail[i].GetValue('PFO_DATEFIN'));
                T.PutValue('NBREHEURE',TobSta.Detail[i].GetValue('PFO_NBREHEURE'));
                T.PutValue('PSA_SALARIE',TobSta.Detail[i].GetValue('PSA_SALARIE'));
        end;
        TobSta.Free;
        TobEtat.Detail.Sort('RUPTURE1;RUPTURE2;RUPTURE3;RUPTURE4');
        StPages := '';   }
        {$IFDEF EAGLCLIENT}
//        StPages := AglGetCriteres (Pages, FALSE);
        {$ENDIF}
//        If GetCheckBoxState('FLISTE') = CbChecked then LanceEtatTOB('E','PFO','PCF',TobEtat,True,True,False,Pages,' ','',False,0,StPages)
//        else LanceEtatTOB('E','PFO','PCF',TobEtat,True,False,False,Pages,' ','',False,0,StPages);
//        TobEtat.Free;
//end;

Procedure TOF_PGEDITCOLLECITVEFORMATION.ConstruireTob ();
begin
        T := Tob.Create('Les filles salariés',TobEtat,-1);
        T.AddChampSup('RUPTURE1',False);T.AddChampSup('RUPTURE2',False);
        T.AddChampSup('RUPTURE3',False);T.AddChampSup('RUPTURE4',False);
        T.AddChampSup('LIBRUPTURE1',False);T.AddChampSup('LIBRUPTURE2',False);
        T.AddChampSup('LIBRUPTURE3',False);T.AddChampSup('LIBRUPTURE4',False);
        T.AddChampSup('TYPE',False);T.AddChampSup('LIBSTAGE',False);
        T.AddChampSup('DATEDEBUT',False);T.AddChampSup('DATEFIN',False);
        T.AddChampSup('NBREHEURE',False);T.AddChampSup('PSA_SALARIE',False);
end;
Function TOF_PGEDITCOLLECITVEFORMATION.RendTablette(Champ : String) : String;
var i : Integer;
begin
Result := '';
If Champ='' then exit;
If Champ = 'PSS_CODESTAGE' then Result := 'PGSTAGEFORM'
else if Champ = 'PSA_SALARIE' then 
Begin
	//PT6 - Début
	If Not PGBundleHierarchie Then
		Result := 'PGSALARIE'
	Else
		Result := 'PGSALARIEINT';
	//PT6 - Fin
End
//PT6
else if Champ = 'PSI_INTERIMAIRE' then Result := 'PGSALARIEINT'
else if Champ = 'PSI_ETABLISSEMENT' then Result := 'TTETABLISSEMENT'
else If Champ = 'PSI_LIBELLEEMPLOI' then Result := 'PGLIBEMPLOI'
//PT6
else if Champ = 'PSA_ETABLISSEMENT' then Result := 'TTETABLISSEMENT'
else If Champ = 'PSA_LIBELLEEMPLOI' then Result := 'PGLIBEMPLOI'
else if Champ = 'PSS_NATUREFORM' then Result := 'PGNATUREFORM'
else if Champ = 'PSS_LIEUFORM' then Result := 'PGLIEUFORMATION'
        else
        begin
        For i :=  1 to VH_Paie.NBFormationLibre do
        begin
                If Champ = 'PSS_FORMATION'+IntToStr(i) then Result := 'PGFORMATION'+IntToStr(i);
        end;
        For i :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                If Champ = 'PSA_TRAVAILN'+IntToStr(i) then Result := 'PGTRAVAILN'+IntToStr(i);
        end;
        end;
end;

{TOF_PGEDITANIMATEURS}
procedure TOF_PGEDITANIMATEURS.OnArgument (S : String ) ;
var Num : Integer;
    Millesime,Combo : THValComboBox;
    DF,DD : TDateTime;
    Check : TCheckBox;
    {$IFDEF EMANAGER}
    Edit : THEdit;
    {$ENDIF}
begin
  Inherited ;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
         VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
        Millesime := THValComboBox(GetControl('MILLESIME'));
        Millesime.Value := RendMillesimeRealise(DD,DF);
        SetControlText('PAN_DATEDEBUT',DateToStr(DD));
        SetcontrolText('PAN_DATEFIN',DateToStr(DF));
        If Millesime <> Nil Then Millesime.OnChange := ChangeMillesime;
        SetControlCaption('LIBSTAGE','');
        SetControlCaption('LIBFORMATEUR','');
        For Num := 1 To 8 do
        begin
                Combo := THValComboBox(GetControl('PST_FORMATION'+IntToStr(Num)));
                If Combo.Visible=True Then SetControlVisible('CRUPT'+IntToStr(Num),True);
                Check := TCheckBox(GetControl('CRUPT'+IntToStr(Num)));
                If Check <> Nil Then Check.OnClick := ChoixRupture;
        end;
        {$IFDEF EMANAGER}
        Edit := THEdit(GetControl('PAN_SALARIE'));
        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        Edit := THEdit(GetControl('PAN_SALARIE_'));
        If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
        {$ENDIF}
end;

procedure TOF_PGEDITANIMATEURS.OnUpdate ;
var Pages : TPageControl;
    Where,OrderBy : String;
begin
  Inherited ;
        Pages := TPageControl(GetControl('Pages'));
        Where := RecupWhereCritere(Pages);
        SetControlText('XX_RUPTURE1','');
        OrderBy := ' ORDER BY PAN_SALARIE,PAN_DATEDEBUT';
        If GetControlText('CRUPT1') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION1,PAN_DATEDEBUT';
                SetControlText('TABLETTE',VH_Paie.FormationLibre1);
                SetControlText('XX_RUPTURE1','PST_FORMATION1');
        end;
        If GetControlText('CRUPT2') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION2,PAN_DATEDEBUT';
                SetControlText('TABLETTE',VH_Paie.FormationLibre2);
                SetControlText('XX_RUPTURE1','PST_FORMATION2');
        end;
        If GetControlText('CRUPT3') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION3,PAN_DATEDEBUT';
                SetControlText('TABLETTE',VH_Paie.FormationLibre3);
                SetControlText('XX_RUPTURE1','PST_FORMATION3');
        end;
        If GetControlText('CRUPT4') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION4,PAN_DATEDEBUT';
                SetControlText('TABLETTE',VH_Paie.FormationLibre4);
                SetControlText('XX_RUPTURE1','PST_FORMATION4');
        end;
        If GetControlText('CRUPT5') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION5,PAN_DATEDEBUT';
                SetControlText('TABLETTE',VH_Paie.FormationLibre5);
                SetControlText('XX_RUPTURE1','PST_FORMATION5');
        end;
        If GetControlText('CRUPT6') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION6,PAN_DATEDEBUT';
                SetControlText('XX_RUPTURE1','PST_FORMATION6');
                SetControlText('TABLETTE',VH_Paie.FormationLibre6);
        end;
        If GetControlText('CRUPT7') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION7,PAN_DATEDEBUT';
                SetControlText('XX_RUPTURE1','PST_FORMATION7');
                SetControlText('TABLETTE',VH_Paie.FormationLibre7);
        end;
        If GetControlText('CRUPT8') = 'X' Then
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PST_FORMATION8,PAN_DATEDEBUT';
                SetControlText('XX_RUPTURE1','PST_FORMATION8');
                SetControlText('TABLETTE',VH_Paie.FormationLibre8);
        end;
        {$IFDEF EMANAGER}
        If Where <> '' then where := ' LEFT JOIN DEPORTSAL ON PAN_SALARIE=PSE_SALARIE '+Where+' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'
        else where := ' LEFT JOIN DEPORTSAL ON PAN_SALARIE=PSE_SALARIE WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        {$ENDIF}
        
        //PT8
	    // N'affiche pas les salariés confidentiels
	    If Where <> '' Then Where := Where + ' AND ';
	    Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
        
        TFQRS1(Ecran).WhereSQL := 'SELECT SESSIONANIMAT.*,PST_LIBELLE,PST_LIBELLE1,'+
        'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5'+
        ',PST_FORMATION6,PST_FORMATION7,PST_FORMATION8 FROM SESSIONANIMAT '+
        'LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE '+
        'LEFT JOIN STAGE ON '+
        'PST_CODESTAGE=PAN_CODESTAGE AND PST_MILLESIME=PAN_MILLESIME '+Where+OrderBy;
end;

procedure TOF_PGEDITANIMATEURS.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEFIN,PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_MILLESIME="'+THValComboBox(Sender).Value+'"',True);
        If not Q.eof Then
        begin
                SetControlText('PAN_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                SetcontrolText('PAN_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        end;
        Ferme(Q);
end;

procedure TOF_PGEDITANIMATEURS.ChoixRupture(Sender : TObject);
var i : Integer;
begin
        If GetCheckBoxState(TCheckBox(Sender).Name)=CbChecked Then
        begin
                For i := 1 to 8 do
                begin
                        If ('CRUPT'+IntToStr(i)) <> TCheckBox(Sender).Name Then SetControlChecked('CRUPT'+IntToStr(i),False);
                end;
        end;
end;

{$IFDEF EMANAGER}
procedure TOF_PGEDITANIMATEURS.SalarieElipsisClick(Sender : TObject);
var StWhere,StOrder : String;
begin
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        StOrder := 'PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
        LookupList(THEdit(Sender),'Liste des stages','SALARIES','PSA_SALARIE LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;
{$ENDIF}

{TOF_PGEDITFORANIMATEUR}
Procedure TOF_PGEDITFORANIMATEUR.OnUpdate ;
var Pages : TPageControl;
    Where,SQL,OrderBy : String;
    {$IFNDEF EMANAGER}
    Select : String;
    AvecLibSession,ForceCodeStage : Boolean;
    i : Integer;
    {$ENDIF}
begin
  Inherited ;

	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS')); //PT6

  OrderBy := '';
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);

  {$IFNDEF EMANAGER}
        For i := 1 to 4 do
        begin
                If GetControlText('VRUPT'+IntToStr(i)) <> '' then
                begin
                        If OrderBy = '' then OrderBy := ' ORDER BY '+GetControlText('VRUPT'+IntToStr(i))
                        else OrderBy := OrderBy+','+GetControlText('VRUPT'+IntToStr(i));
                        SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
                end
                else SetControlText('XX_RUPTURE'+IntToStr(i),'');
        end;
        If LeMenu = '47541' then     //Suivi individuel des animateurs
        begin
                OrderBy := ' ORDER BY PAN_SALARIE,PAN_CODESTAGE,PAN_ORDRE';
                TFQRS1(Ecran).CodeEtat := 'PAI';
                TFQRS1(Ecran).FCodeEtat := 'PAI';
                
                //PT6
                SQL := 'SELECT PAN_SALARIE,PAN_CODESTAGE,PAN_ORDRE,PSS_CODESTAGE,PSS_ORDRE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_NATUREFORM,PSS_DUREESTAGE,PSS_CENTREFORMGU,'+
                       'PSS_FORMATION1,PSS_FORMATION2,PSS_FORMATION3,PSS_FORMATION4,PSS_FORMATION5,PSS_FORMATION6,PSS_FORMATION7,PSS_FORMATION8,'+
              		   'PSA_ETABLISSEMENT AS ETABLISSEMENT,PSA_LIBELLEEMPLOI AS LIBELLEEMPLOI,PSA_TRAVAILN1 AS TRAVAILN1,PSA_TRAVAILN2 AS TRAVAILN2,PSA_TRAVAILN3 AS TRAVAILN3,PSA_TRAVAILN4 AS TRAVAILN4,PSA_DATEENTREE AS DATEENTREE'; 
                	
                SQL := SQL + ' FROM SESSIONANIMAT '+
                             ' LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME ';
                
                If PGBundleHierarchie Then //PT6
                	SQL := SQL + ' LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PAN_SALARIE '
                Else
                	SQL := SQL + ' LEFT JOIN SALARIES ON PSA_SALARIE=PAN_SALARIE ';
			    
			    //PT8
			    // N'affiche pas les salariés confidentiels
			    If Where <> '' Then Where := Where + ' AND ';
			    Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
			    
				If PGBundleHierarchie Then //PT6 
				Begin
                    SQL   := StringReplace (SQL,   'PSA_', 'PSI_', [rfIgnoreCase,rfReplaceAll]);
					Where := StringReplace (Where, 'PSA_', 'PSI_', [rfIgnoreCase,rfReplaceAll]);
                	Where := StringReplace (Where, 'PSS_NODOSSIER', 'PSI_NODOSSIER', [rfReplaceAll,rfIgnoreCase]);
                End;
                //SQL := 'SELECT SESSIONANIMAT.*,PSA_DATEENTREE AS DATEENTREE FROM SESSIONANIMAT LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE ';
                //SQL := 'SELECT * FROM SESSIONANIMAT LEFT JOIN SALARIES ON PSA_SALARIE=PAN_SALARIE '+
                
                //PT6 - Fin
                TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
        end;
        If LeMenu = '47542' then     //Suivi collectif des animateurs
        begin
                TFQRS1(Ecran).CodeEtat := 'PAC';
                TFQRS1(Ecran).FCodeEtat := 'PAC';
                SQL := 'SELECT * FROM SESSIONANIMAT '+
                'LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME ';
                
                If Where <> '' then where := Where + ' AND PAN_SALARIE<>""'
                Else Where := 'WHERE PAN_SALARIE<>""';
			    
			    //PT8
			    // N'affiche pas les salariés confidentiels
			    Where := Where + ' AND PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
                
                //PT6 - Début
                If PGBundleHierarchie Then
                Begin
                	SQL := SQL + ' LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PAN_SALARIE ';
                	Where := StringReplace(Where, 'PSA_', 'PSI_', [rfIgnoreCase,rfReplaceAll]);
                End
                Else
                	SQL := SQL + ' LEFT JOIN SALARIES ON PSA_SALARIE=PAN_SALARIE ';
                //PT6 - Fin
                
                TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
        end;
        If LeMenu = '47562' then  //Frais animateurs
        begin
                TFQRS1(Ecran).CodeEtat := 'PFA';
                TFQRS1(Ecran).FCodeEtat := 'PFA';
                ForceCodeStage := True;
                AvecLibSession := False;
                For i := 1 to 4 do
                begin
                        SetControlText('XX_RUPTURE'+IntToStr(i),GetControlText('VRUPT'+IntToStr(i)));
                        If GetControlText('VRUPT'+IntToStr(i)) <> '' then
                        begin
                                If GetControlText('VRUPT'+IntToStr(i)) = 'PAN_ORDRE' then  AvecLibSession := True;
                                If OrderBy <> '' then OrderBy := OrderBy + ',' + GetControlText('VRUPT'+IntToStr(i))
                                else OrderBy := ' ORDER BY '+ GetControlText('VRUPT'+IntToStr(i));
                                If (GetControlText('VRUPT'+IntToStr(i)) <> 'PAN_CODESTAGE') and (GetControlText('VRUPT'+IntToStr(i)) <> 'PAN_ORDRE') then
                                Select := Select+','+GetControlText('VRUPT'+IntToStr(i));

                        end;
                end;
                If AvecLibSession = True then
                begin
                        For i := 1 to 4 do
                        begin
                                If GetControlText('VRUPT'+IntToStr(i)) = 'PAN_CODESTAGE' then ForceCodeStage := False;
                        end;
                end
                else forceCodeStage := False;
                If ForceCodeStage = True then
                begin
                        If OrderBy <> '' then OrderBy := OrderBy + ',PAN_CODESTAGE'
                        else OrderBy := ' ORDER BY PAN_CODESTAGE';
                end;
                
			    //PT8
			    // N'affiche pas les salariés confidentiels
			    If Where <> '' Then Where := Where + ' AND ';
			    Where := Where + 'PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"';
			    
                TFQRS1(Ecran).WhereSQL := 'SELECT PAN_SALARIE,PAN_CODESTAGE,PAN_ORDRE,PAN_MILLESIME,PAN_DATEDEBUT,PAN_DATEFIN,PSS_LIBELLE'+Select+' FROM SESSIONANIMAT '+
                                          //'LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE '+ //PT6
                                          'LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME '+Where+OrderBy;
				//PT6 - Début				                                          
				If PGBundleHierarchie Then
				Begin
					TFQRS1(Ecran).WhereSQL := StringReplace(TFQRS1(Ecran).WhereSQL, 'PSA_', 'PSI_', [rfIgnoreCase,rfReplaceAll]);
					TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL  + 'LEFT JOIN INTERIMAIRES ON PAN_SALARIE=PSI_INTERIMAIRE ';
				End
				Else
					TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL  + 'LEFT JOIN SALARIES ON PAN_SALARIE=PSA_SALARIE ';
				//PT6 - Fin
        end;
     {$ELSE}
     //PT4 - Début
     If LeMenu = 'COLLECTIF' then
     Begin
          TFQRS1(Ecran).CodeEtat := 'PAC';
          TFQRS1(Ecran).FCodeEtat := 'PAC';
          TFQRS1(Ecran).NatureEtat := 'PFE';
          
          SQL := 'SELECT * FROM SESSIONANIMAT '+
          ' LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME ';
          
          If Where <> '' then where := Where + ' AND PAN_SALARIE<>""'
          Else Where := 'WHERE PAN_SALARIE<>""';
          
          If TypeUtilisat = 'R' Then
               Where := Where + ' AND PAN_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")'
          Else
               Where := Where + ' AND PAN_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))';
					
          TFQRS1(Ecran).WhereSQL := SQL+Where+OrderBy;
     End;
     //PT4 - Fin
     {$ENDIF}
end ;

{$IFNDEF EMANAGER}
procedure TOF_PGEDITFORANIMATEUR.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;
{$ENDIF}

procedure TOF_PGEDITFORANIMATEUR.OnArgument (S : String ) ;
var {$IFNDEF EMANAGER}
    Num,i : Integer;
    Combo : THValComboBox;
    {$ENDIF}
    DD,DF : TDateTime;
    Edit : THEdit;
    Prefixe : String; //PT6
begin
  Inherited ;
     {$IFNDEF EMANAGER} //PT4 //PT6
     Edit := THEdit(GetControl('PAN_SALARIE'));
     If Edit <> Nil then
     Begin
        Edit.OnExit := ExitEdit;
        Edit.OnElipsisClick := SalarieElipsisClick; //PT6
     End;
     Edit := THEdit(GetControl('PAN_SALARIE_'));
     If Edit <> Nil then
     Begin
        Edit.OnExit := ExitEdit;
        Edit.OnElipsisClick := SalarieElipsisClick; //PT6
     End;
     {$ENDIF}


     //PT6 - Début
     If PGBundleHierarchie Then
        Prefixe := 'PSI'
     Else
        Prefixe := 'PSA';
     //PT6 - Fin

     LeMenu := ReadTokenPipe(S,';');
     SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
     RendMillesimeRealise(DD,DF);
     SetControlText('PSS_DATEDEBUT',DateToStr(DD));
     SetControlText('PSS_DATEFIN',DateToStr(DF));

     {$IFNDEF EMANAGER}
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PSS_FORMATION'+IntToStr(Num)),GetControl ('TPSS_FORMATION'+IntToStr(Num)));
        end;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
                SetControlVisible('CRUPTSAL'+IntToStr(Num),True);
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
        If LeMenu = '47562' then  // Frais des animateurs
        begin
                SetControlVisible('SANSDETAIL',False);
                Ecran.Caption := 'Edition des frais par animateur';
        end;
        If LeMenu = '47541' then     //Suivi pédagogique stagiaire
        begin
                SetControlVisible('SANSDETAIL',False);
                SetControlProperty('PRUPT','TabVisible',False);
                Ecran.Caption := 'Suivi animation interne individuel';
        end;
        If LeMenu = '47542' then     //Suivi pédagogique stagiaire
        begin

                Ecran.Caption := 'Suivi animation interne collectif';
        end;
        UpdateCaption(Ecran);
        For i := 1 To 4 do
        begin
                Combo := THValComboBox(GetControl('VRUPT'+IntToStr(i)));
                If Combo <> Nil Then
                begin
                       // If i <> 1 Then
                       // begin
                                Combo.Items.Add('<<Aucune>>');
                                Combo.Values.Add('');
                       // end;
                        Combo.Items.Add('Salarié');
                        Combo.Values.Add('PAN_SALARIE');
                        Combo.Items.Add('Formation');
                        Combo.Values.Add('PAN_CODESTAGE');
                        Combo.Items.Add('Etablissement');
                       	Combo.Values.Add(Prefixe+'_ETABLISSEMENT');
                        Combo.Items.Add('Libellé emploi');
                       	Combo.Values.Add(Prefixe+'_LIBELLEEMPLOI');
                        Combo.Items.Add('Nature (Interne/Externe)');
                        Combo.Values.Add('PSS_NATUREFORM');
                        Combo.Items.Add('Lieu');
                        Combo.Values.Add('PSS_LIEUFORM');
                        Combo.Items.Add('Centre de formation');
                        Combo.Values.Add('PSS_CENTREFORMGU');
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
                        If VH_Paie.PGLibCodeStat <> '' Then begin Combo.Items.Add(VH_Paie.PGLibCodeStat); Combo.Values.Add('PSA_CODESTAT');end;
                end;
        end;
     {$ELSE}
     //PT4 - Début
     If LeMenu = 'COLLECTIF' then Ecran.Caption := 'Suivi animation interne collectif';
     UpdateCaption(Ecran);

     If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
     else TypeUtilisat := 'S';

     SetControlProperty('PSA_LIBELLEEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
     'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsdateTime(V_PGI.DateEntree)+'" OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'") AND '+
     AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,False)+')');

	 If PGBundleHierarchie Then //PT7
	 Begin
	     Edit := THEdit(GetControl('PAN_SALARIE'));
	     if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
	 End;
     //PT4 - Fin
     {$ENDIF}
     
     
	//PT6 - Début
	If PGBundleCatalogue Then //PT7
	Begin
		If not PGDroitMultiForm then
			SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
       		SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PSS_PREDEFINI',False);
			SetControlText   ('PSS_PREDEFINI','');
			//SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT6 //PT7
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT6
       		//SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6 //PT7
       	End;
	end
	else
	begin
		SetControlVisible('PSS_PREDEFINI', False);
		SetControlVisible('TPSS_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT6 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 26/10/2007 / PT4
Modifié le ... :   /  /
Description .. : Elipsis pour salariés rattachés au manager
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITFORANIMATEUR.SalarieElipsisClick(Sender: TObject);
{$IFDEF EMANAGER}
var
  StFrom, StWhere: string;
{$ENDIF}
begin
{$IFDEF EMANAGER}
  StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  StWhere := StWhere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked));

  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
{$ELSE}
     //PT6 - Début
    //If PGBundleInscFormation Then //PT7
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
{$ENDIF}
end;


{TOF_EDITSTAGIAIREANIMATEUR}
Procedure TOF_EDITSTAGIAIREANIMATEUR.OnUpdate ;
var Pages : TPageControl;
    Where,OrderBy : String;
    EtatChoisi : String;
    Req : String; //PT6
begin
  Inherited ;
  
	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS')); //PT6	
	
	OrderBy := '';
	Pages := TPageControl(GetControl('Pages'));
	Where := RecupWhereCritere(Pages);
	EtatChoisi := GetControlText('FETAT');

	//PT6 - Début
	If PGBundleInscFormation Then
	Begin
		If LeMenu = '47523' Then
		Begin
			Req :=  'SELECT PAN_SALARIE,PFO_CODESTAGE,PAN_NBREHEURE,PFO_ORDRE,PSS_LIEUFORM,PAN_SALAIREANIM,PSS_LIBELLE,'+
					'PFO_DATEDEBUT,PFO_DATEFIN,PFO_AUTRECOUT,PFO_COUTREELSAL,PFO_FRAISREEL,PFO_SALARIE,PFO_NBREHEURE,'+
					'"01/01/1900" AS ANCANIM,S1.PSI_DATENAISSANCE AS NAISSANIM,S1.PSI_ETABLISSEMENT AS ETABLISSEMENT,S1.PSI_LIBELLEEMPLOI AS LIBELLEEMPLOI,S1.PSI_TRAVAILN1 AS TRAVAILN1,'+
					'"01/01/1900" AS ANCSTA,S2.PSI_DATENAISSANCE AS NAISSSTA,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_TRAVAILN1,PFO_TRAVAILN4,'+
					'S2.PSI_LIBREPCMB2 AS LIBREPCMB2 '+
					'FROM FORMATIONS '+
					'LEFT JOIN SESSIONSTAGE ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME '+
					'LEFT JOIN INTERIMAIRES S2 ON PFO_SALARIE=S2.PSI_INTERIMAIRE '+
					'LEFT JOIN SESSIONANIMAT ON PFO_CODESTAGE=PAN_CODESTAGE AND PFO_ORDRE=PAN_ORDRE AND PFO_MILLESIME=PAN_MILLESIME '+
					'LEFT JOIN INTERIMAIRES S1 ON PAN_SALARIE=S1.PSI_INTERIMAIRE ';
			If Where <> '' Then
				Req := Req + Where + ' AND '
			Else
				Req := Req + ' WHERE ';
			Req := Req + ' PFO_EFFECTUE="X" AND PAN_SALARIE<>"" '+
						 'ORDER BY PFO_SALARIE,PFO_CODESTAGE,PFO_ORDRE';
		End
		Else If LeMenu = '47522' Then
		Begin
			Req :=  'SELECT PAN_SALARIE,PAN_CODESTAGE,PAN_NBREHEURE,PAN_ORDRE,PSS_LIEUFORM,PSS_LIBELLE,'+
					'PFO_DATEDEBUT,PFO_DATEFIN,PFO_AUTRECOUT,PFO_COUTREELSAL,PFO_FRAISREEL,PFO_SALARIE,PFO_NBREHEURE,'+
					'"01/01/1900" AS ANCANIM,S1.PSI_DATENAISSANCE AS NAISSANIM,S1.PSI_ETABLISSEMENT AS ETABLISSEMENT,S1.PSI_LIBELLEEMPLOI AS LIBELLEEMPLOI,S1.PSI_TRAVAILN1 AS TRAVAILN1,'+
					'S1.PSI_TRAVAILN4 AS TRAVAILN4,S1.PSI_LIBREPCMB2 AS LIBREPCMB2,'+
					'"01/01/1900" ANCSTA,S2.PSI_DATENAISSANCE NAISSTA,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR,PFO_TRAVAILN1 '+
					'FROM SESSIONANIMAT '+
					'LEFT JOIN SESSIONSTAGE ON PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_MILLESIME=PSS_MILLESIME '+
					'LEFT JOIN INTERIMAIRES S1 ON PAN_SALARIE=S1.PSI_INTERIMAIRE '+
					'LEFT JOIN FORMATIONS ON PFO_CODESTAGE=PAN_CODESTAGE AND PFO_ORDRE=PAN_ORDRE AND PFO_MILLESIME=PAN_MILLESIME '+
					'LEFT JOIN INTERIMAIRES S2 ON PFO_SALARIE=S2.PSI_INTERIMAIRE ';
			If Where <> '' Then 
				Req := Req + Where + ' AND ' 
			Else 
				Req := Req + ' WHERE ';
			Req := Req + ' PFO_EFFECTUE="X" AND PAN_SALARIE<>"" '+
						 'ORDER BY PAN_SALARIE,PAN_CODESTAGE,PAN_ORDRE';
		End;
		TFQRS1(Ecran).WhereSQL := Req;
	End;
	//PT6 - Fin

	if (EtatChoisi <> 'PAD') OR (EtatChoisi <> 'PAE') then Exit; //PT2
	If LeMenu = '47522' then
	begin
		TFQRS1(Ecran).CodeEtat := 'PAB';
		TFQRS1(Ecran).FCodeEtat := 'PAB';
	end;
	If LeMenu = '47523' then
	begin
		TFQRS1(Ecran).CodeEtat := 'PAA';
		TFQRS1(Ecran).FCodeEtat := 'PAA';
	end;
end ;

procedure TOF_EDITSTAGIAIREANIMATEUR.SalarieElipsisClick(Sender : TObject);
begin
     //PT6 - Début
    //If PGBundleInscFormation Then //PT7
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
end;

procedure TOF_EDITSTAGIAIREANIMATEUR.OnArgument (S : String ) ;
var Num : Integer;
    DF,DD : TDateTime;
    Edit  : THEdit; //PT6
begin
  Inherited ;
	LeMenu := ReadTokenPipe(S,';');
	SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
	RendMillesimeRealise(DD,DF);
	SetControlText('PSS_DATEDEBUT',DateToStr(DD));
	SetControlText('PSS_DATEDEBUT_',DateToStr(DF));
	For Num  :=  1 to VH_Paie.PGNbreStatOrg do
	begin
		if Num >4 then Break;
		VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
	end;
	VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
	If LeMenu = '47522' then
	begin
		Ecran.Caption := 'Suivi stagiaires / animateurs';
	end;
	If LeMenu = '47523' then
	begin
		Ecran.Caption := 'Suivi animateurs / stagiaires';
	end;
	UpdateCaption(Ecran);
	If LeMenu = '47522' then
	begin
		TFQRS1(Ecran).CodeEtat := 'PAB';
		TFQRS1(Ecran).FCodeEtat := 'PAB';
	end;
	If LeMenu = '47523' then
	begin
		TFQRS1(Ecran).CodeEtat := 'PAA';
		TFQRS1(Ecran).FCodeEtat := 'PAA';
	end;
	
    //PT8
    // N'affiche pas les salariés confidentiels
    If PGBundleHierarchie Then
        SetControlText('XX_WHERE', '(S1.PSI_CONFIDENTIEL IS NULL OR S1.PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'") AND (S1.PSI_CONFIDENTIEL IS NULL OR S2.PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")')  //PT9
    Else
        SetControlText('XX_WHERE', '(S1.PSA_CONFIDENTIEL IS NULL OR S1.PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'") AND (S1.PSA_CONFIDENTIEL IS NULL OR S2.PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")');  //PT9

    //PT6 - Début
    If PGBundleCatalogue Then //PT7
    Begin
    	If not PGDroitMultiForm then
			SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
       		SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PSS_PREDEFINI',False);
			SetControlText   ('PSS_PREDEFINI','');
			//SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT6 //PT7
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT6
       		//SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6 //PT7
       	End;
	end
	else
	begin
		SetControlVisible('PSS_PREDEFINI', False);
		SetControlVisible('TPSS_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	
	Edit := THEdit(GetControl('PFO_SALARIE'));
	If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	Edit := THEdit(GetControl('PFO_SALARIE_'));
	If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	Edit := THEdit(GetControl('PAN_SALARIE'));
	If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	Edit := THEdit(GetControl('PAN_SALARIE_'));
	If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	//PT6 - Fin
end ;


Initialization
registerclasses([TOF_PGEDITSALARIESFORMATION,TOF_PGEDITCOLLECITVEFORMATION,TOF_PGEDITANIMATEURS,TOF_PGEDITFORANIMATEUR,TOF_EDITSTAGIAIREANIMATEUR]);
end.

