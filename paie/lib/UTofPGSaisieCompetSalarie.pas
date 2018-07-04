{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 11/02/2004
Modifié le ... :   /  /
Description .. : Source TOF des fiches de saisie de compétences par : salarié,intérimaire,candidat
Mots clefs ... : TOF;PGSAISIECOMPETSALARIE
*****************************************************************
PT1 24/11/2004 V_60 JL Correction LookUpList
}
Unit UTofPGSaisieCompetSalarie ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,
{$else}
     eMul,uTob,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,P5Def,EntPaie,SaisieList,
     uTableFiltre,PGOutilsFormation,HTB97,LookUp ,PGEditOutils,PGOutils, PgOutils2;

Type
  TOF_PGSAISIECOMPETSALARIE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    NatureSaisie : String;
    TF : TTableFiltre;
    procedure OnClickSalarieSortie(Sender: TObject);
    procedure AffichageNomCaption(Sender : TObject);
    procedure ChangerPeriode(Sender : TObject);
    procedure CompetenceElipsisClick(Sender : TObject);
    procedure ParamTV(Sender : TObject);
    procedure ExitEdit(Sender: TObject);
    procedure AccesColonne(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
  end ;

Type
  TOF_PGSAISIECOMPETSTAGE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    NatureSaisie : String;
    TF : TTableFiltre;
    procedure AffichageNomCaption(Sender : TObject);
    procedure ExitEdit(Sender: TObject);
  end ;

Implementation

procedure TOF_PGSAISIECOMPETSALARIE.AccesColonne(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var i : Integer;
begin
        For i := 1 to 5 do
        begin
                If TF.LaGrid.ColNames[Ou] = 'PCO_TABLELIBRERH' + IntToStr(i) then
                begin
                        If Ou < (TF.LaGrid.ColCount - 1) then TF.LaGrid.Col := Ou + 1
                        else TF.LaGrid.Col := 1;
                end;
        end;
end;

procedure TOF_PGSAISIECOMPETSALARIE.ParamTV(Sender : TObject);
var DateArret : TDateTime;
    StDateArret,NomCol,FiltreTable : String;
    i : Integer;
begin
        DateArret := StrToDate(GetControlText('DATEPRESENCE'));
        If (NatureSaisie = 'SALARIES') or (NatureSaisie = 'INTERIMAIRES') then
        begin
                if (GetControlText('CPRESENT')='X') and (IsValidDate(GetControlText('DATEPRESENCE')))then
                Begin
                        StDateArret:=' AND (PSI_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSI_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSI_DATESORTIE IS NULL) ';
                        StDateArret:=StDateArret + ' AND PSI_DATEENTREE <="'+UsDateTime(DateArret)+'"';
                        SetControlText('XX_WHERE',StDateArret);
                End
                else SetControlText('XX_WHERE','');
                FiltreTable := '';
                TFSaisieList(Ecran).RichMessage1.Clear;
                For i := 1 to 5 do
                begin
                        If GetControlText('TABLELIBRERH'+IntToStr(i)) <> '' then
                        begin
                                FiltreTable := FiltreTable + ' AND PCO_TABLELIBRERH'+IntToStr(i)+'="'+GetControlText('TABLELIBRERH'+IntToStr(i))+'"';
                                TFSaisieList(Ecran).RichMessage1.Lines.Append('Libre '+IntToStr(i)+' = '+GetControlText('TABLELIBRERH'+IntToStr(i)));
                        end;
                end;
                TF.WhereTable:='WHERE PCH_TYPERESSOURCE=:PSI_TYPEINTERIM AND PCH_SALARIE=:PSI_INTERIMAIRE AND PCH_DATEVALIDATION>="'+UsDateTime(StrToDate(GetControlText('DATE1')))+'" AND '+
                'PCH_DATEVALIDATION<="'+UsDateTime(StrToDate(GetControlText('DATE2')))+'"'+FiltreTable;
                TF.RefreshLignes;
        end;
        TFSaisieList(Ecran).BChercheClick(Sender);
end;

procedure TOF_PGSAISIECOMPETSALARIE.OnLoad ;
var DateArret : TDateTime;
    StDateArret,NomCol,FiltreTable,Libelle : String;
    i,x : Integer;
begin
  Inherited ;

//TF.ChangeColFormat('PCO_TABLELIBRERH2','CB=PGTABLELIBRERH2');
//TF.ChangeColFormat('PCO_TABLELIBRERH3','CB=PGTABLELIBRERH3');
//TF.ChangeColFormat('PCO_TABLELIBRERH4','CB=PGTABLELIBRERH4');
//TF.ChangeColFormat('PCO_TABLELIBRERH5','CB=PGTABLELIBRERH5');
For i := 1 to TF.LaGrid.ColCount - 1 do
begin
        NomCol := TF.LaGrid.ColNames[i];
        If NomCol = 'PCO_LIBELLE' then TF.LaGrid.ColEditables[i]:= False;
        If NomCol = 'PCO_TABLELIBRERH1' then
        begin
//                If TF.Reccount > 0 then
//                begin
//                        For x := 0 To TF.TobFiltre.Detail.Count - 1 do
//                        begin
//                                Libelle := TF.TOBFiltre.Detail[x].GetValue('PCO_TABLELIBRERH1');
//                                TF.TOBFiltre.Detail[x].PutValue('PCO_TABLELIBRERH1',RechDom('PGTABLELIBRERH1',Libelle,False));
//                        end;
//                end;
                TF.ChangeColFormat('PCO_TABLELIBRERH1','CB=PGTABLELIBRERH1');
                TF.LaGrid.ColEditables[i]:= False;
        end;
        If NomCol = 'DEGREMAITRISEEMP' then TF.LaGrid.ColEditables[i]:= False;
end;
SetControlEnabled('PCO_TABLELIBRERH2',False);
DateArret := StrToDate(GetControlText('DATEPRESENCE'));
If (NatureSaisie = 'SALARIES') or (NatureSaisie = 'INTERIMAIRES') then
begin
  if (GetControlText('CPRESENT')='X') and (IsValidDate(GetControlText('DATEPRESENCE')))then
     Begin
     StDateArret:=' AND (PSI_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSI_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSI_DATESORTIE IS NULL) ';
     StDateArret:=StDateArret + ' AND PSI_DATEENTREE <="'+UsDateTime(DateArret)+'"';
     SetControlText('XX_WHERE',StDateArret);
     End
  else SetControlText('XX_WHERE','');
  FiltreTable := '';
  TFSaisieList(Ecran).RichMessage1.Clear;
  For i := 1 to 5 do
  begin
          If GetControlText('TABLELIBRERH'+IntToStr(i)) <> '' then
          begin
                FiltreTable := FiltreTable + ' AND PCO_TABLELIBRERH'+IntToStr(i)+'="'+GetControlText('TABLELIBRERH'+IntToStr(i))+'"';
                TFSaisieList(Ecran).RichMessage1.Lines.Append('Libre '+IntToStr(i)+' = '+GetControlText('TABLELIBRERH'+IntToStr(i)));
          end;
  end;
  TF.WhereTable:='WHERE PCH_TYPERESSOURCE=:PSI_TYPEINTERIM AND PCH_SALARIE=:PSI_INTERIMAIRE AND PCH_DATEVALIDATION>="'+UsDateTime(StrToDate(GetControlText('DATE1')))+'" AND '+
        'PCH_DATEVALIDATION<="'+UsDateTime(StrToDate(GetControlText('DATE2')))+'"'+FiltreTable;
  TF.RefreshLignes;
end;
If NatureSaisie = 'GESTCARRIERESAL' then
begin
        SetControlVisible('BTree',False);
        SetControlVisible('PANTREEVIEW',False);
end;
TFSaisieList(Ecran).BCherche.Click;
end ;

procedure TOF_PGSAISIECOMPETSALARIE.ChangerPeriode(Sender : TObject);
var i : integer;
    FiltreTable : String;
begin
        For i := 1 to 5 do
        begin
          If GetControlText('TABLELIBRERH'+IntToStr(i)) <> '' then
            FiltreTable := FiltreTable + ' AND PCO_TABLELIBRERH'+IntToStr(i)+'="'+GetControlText('TABLELIBRERH'+IntToStr(i))+'"';
        end;
        TF.WhereTable:='WHERE PCH_TYPERESSOURCE=:PSI_TYPEINTERIM AND PCH_SALARIE=:PSI_INTERIMAIRE AND PCH_DATEVALIDATION>="'+UsDateTime(StrToDate(GetControlText('DATE1')))+'" AND '+
        'PCH_DATEVALIDATION<="'+UsDateTime(StrToDate(GetControlText('DATE2')))+'"'+FiltreTable;
        TF.RefreshLignes;
end;

procedure TOF_PGSAISIECOMPETSALARIE.OnArgument (S : String ) ;
var Num : Integer;
    Check : TCheckBox;
    Salarie : String;
    BPeriode : TToolBarButton97;
    Edit : THEdit;
begin
  Inherited ;
TF  :=  TFSaisieList(Ecran).LeFiltre;
TF.LaGrid.OnColEnter := AccesColonne;
//TF.LaGrid.OnCellEnter
TFSaisieList(Ecran).BCherche.OnClick := ParamTV;
TF.WhereTable:='WHERE PCH_TYPERESSOURCE=:PSI_TYPEINTERIM AND PCH_SALARIE=:PSI_INTERIMAIRE AND PCH_DATEVALIDATION>="'+UsDateTime(StrToDate(GetControlText('DATE1')))+'" AND '+
        'PCH_DATEVALIDATION<="'+UsDateTime(StrToDate(GetControlText('DATE2')))+'"';
 BPeriode := TToolBarButton97(GetControl('BAPPLIQUER'));
 If Bperiode <> Nil then BPeriode.OnClick := ChangerPeriode;
NatureSaisie := ReadTokenPipe(S,';');
Salarie := ReadTokenPipe(S,';');
SetControlText('XX_WHERE','');
For Num := 1 to VH_Paie.PGNbreStatOrg do
begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSI_TRAVAILN'+IntToStr(Num)),GetControl ('TPSI_TRAVAILN'+IntToStr(Num)));
end;
VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
For Num  :=  1 to 3 do
        begin
                VisibiliteChampCompetenceRess (IntToStr(Num),GetControl ('PCH_TABLELIBRECR'+IntToStr(Num)),GetControl ('TPCH_TABLELIBRECR'+IntToStr(Num)));
        end;
If NatureSaisie = 'SALARIES' then
begin
        Edit := THEdit(GetControl('PCH_COMPETENCE'));
        If Edit <> Nil then Edit.OnElipsisClick := CompetenceElipsisClick;
        TFSaisieList(Ecran).Caption := 'Saisie des compétences par salariés';
        For Num  :=  1 to 5 do
        begin
                VisibiliteChampCompetence (IntToStr(Num),GetControl ('TABLELIBRERH'+IntToStr(Num)),GetControl ('TTABLELIBRERH'+IntToStr(Num)));
        end;
end;
If NatureSaisie = 'CANDIDATS' then
begin
        TFSaisieList(Ecran).Caption := 'Saisie des compétences par candidats';
        SetControlVisible('CPRESENT',False);
        SetControlVisible('DATEPRESENCE',False);
end;
If NatureSaisie = 'INTERIMAIRES' then
begin
        TFSaisieList(Ecran).Caption := 'Saisie des compétences par intérimaires/stagiaires';
        SetControlProperty('PSI_TYPEINTERIM','Plus','AND (CO_CODE="STA" OR CO_CODE="INT")');
end;
If NatureSaisie = 'GESTCARRIERESAL' then
begin
        SetControlText('PSI_INTERIMAIRE',Salarie);
        TFSaisieList(Ecran).Caption := 'Saisie des compétences pour le salarié '+Salarie+' '+RechDom('PGSALARIE',Salarie,False);
end;
UpdateCaption(TFSaisieList(Ecran));
Check := TCheckBox(GetControl('CPRESENT'));
If Check <> Nil then Check.OnClick:=OnClickSalarieSortie;
SetControlEnabled('DATEPRESENCE',False);
SetControlText('DATEPRESENCE',DateToStr(Date));
TF.OnSetNavigate  :=  AffichageNomCaption;
Edit := THEdit(GetControl('PSI_INTERIMAIRE'));
If Edit <> Nil then Edit.OnExit := ExitEdit;
end ;

procedure TOF_PGSAISIECOMPETSALARIE.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEPRESENCE',(GetControltext('CPRESENT')='X'));
end;

procedure TOF_PGSAISIECOMPETSALARIE.AffichageNomCaption(Sender : TObject);
var StNom : String;
begin
StNom := TF.TobFiltre.GetValue('PSI_INTERIMAIRE') + ' ' + TF.TobFiltre.GetValue('PSI_LIBELLE') + ' ' + TF.TobFiltre.GetValue('PSI_PRENOM');
If NatureSaisie = 'SALARIES' then TFSaisieList(Ecran).Caption := 'Saisie des compétences pour le salarié : '+StNom;
If NatureSaisie = 'CANDIDATS' then TFSaisieList(Ecran).Caption := 'Saisie des compétences pour le candidat : '+StNom;
If NatureSaisie = 'INTERIMAIRES' then
begin
        If TF.TobFiltre.GetValue('PSI_TYPEINTERIM') = 'INT' then
        TFSaisieList(Ecran).Caption := 'Saisie des compétences pour l''intérimaire : '+StNom
        else TFSaisieList(Ecran).Caption := 'Saisie des compétences pour le stagiaire : '+StNom;
end;
UpdateCaption(TFSaisieList(Ecran));
end;

procedure TOF_PGSAISIECOMPETSALARIE.CompetenceElipsisClick(Sender : TObject);
var St,StWhere,StOrder : String;
begin
       St := 'RHCOMPETENCES '+
        'LEFT JOIN CHOIXEXT C1 ON C1.YX_CODE=PCO_TABLELIBRERH1 AND C1.YX_TYPE="PR1" ';
{        'LEFT JOIN CHOIXEXT C2 ON C2.YX_CODE=PCO_TABLELIBRERH2 AND C2.YX_TYPE="PR2" '+
        'LEFT JOIN CHOIXEXT C3 ON C3.YX_CODE=PCO_TABLELIBRERH3 AND C3.YX_TYPE="PR3" '+
        'LEFT JOIN CHOIXEXT C4 ON C4.YX_CODE=PCO_TABLELIBRERH4 AND C4.YX_TYPE="PR4" '+
        'LEFT JOIN CHOIXEXT C5 ON C5.YX_CODE=PCO_TABLELIBRERH5 AND C5.YX_TYPE="PR5"';}
        StWhere := '';
        StOrder := 'PCO_COMPETENCE';
//        LookupList(THEdit(Sender),'Liste des compétences',St,'PCO_COMPETENCE','PCO_LIBELLE,C1.YX_LIBELLE LIBRE1,C2.YX_LIBELLE LIBRE2,C3.YX_LIBELLE LIBRE3,C4.YX_LIBELLE LIBRE4,C5.YX_LIBELLE LIBRE5',StWhere,StOrder, True,-1);
        LookupList(THEdit(Sender),'Liste des compétences',St,'PCO_COMPETENCE','PCO_LIBELLE,C1.YX_LIBELLE LIBRE1',StWhere,StOrder, True,-1);  //PT1
end;

procedure TOF_PGSAISIECOMPETSALARIE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
        edit:=THEdit(Sender);
        if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;


//TOF_PGSAISIECOMPETSTAGE

procedure TOF_PGSAISIECOMPETSTAGE.OnLoad ;
begin
  Inherited ;

end ;

procedure TOF_PGSAISIECOMPETSTAGE.OnArgument (S : String ) ;
var Num : Integer;
    Edit : THEdit;
begin
  Inherited ;
NatureSaisie := ReadTokenPipe(S,';');
TF  :=  TFSaisieList(Ecran).LeFiltre;
TF.OnSetNavigate  :=  AffichageNomCaption;
If NatureSaisie = 'STAGE' then
begin
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
end;
Edit := THEdit(GetControl('PST_CODESTAGE'));
If Edit <> Nil then Edit.OnExit := ExitEdit;
end;

procedure TOF_PGSAISIECOMPETSTAGE.AffichageNomCaption(Sender : TObject);
var StNom : String;
begin
If NatureSaisie = 'EMPLOI' then
begin
        StNom := TF.TobFiltre.GetValue('CC_CODE') + ' ' + TF.TobFiltre.GetValue('CC_LIBELLE');
        TFSaisieList(Ecran).Caption := 'Saisie des compétences pour le libellé emploi : '+StNom;
end;
If NatureSaisie = 'STAGE' then
begin
        StNom := TF.TobFiltre.GetValue('PST_CODESTAGE') + ' ' + TF.TobFiltre.GetValue('PST_LIBELLE') + ' ' + TF.TobFiltre.GetValue('PST_LIBELLE1');
        TFSaisieList(Ecran).Caption := 'Saisie des compétences pour la formation : '+StNom;
end;
UpdateCaption(TFSaisieList(Ecran));
end ;

procedure TOF_PGSAISIECOMPETSTAGE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
        edit:=THEdit(Sender);
        if edit <> nil then	//AffectDefautCode que si gestion du code stage en numérique
        if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<6) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,6);
end;


Initialization
  registerclasses ( [ TOF_PGSAISIECOMPETSALARIE,TOF_PGSAISIECOMPETSTAGE ] ) ;
end.

