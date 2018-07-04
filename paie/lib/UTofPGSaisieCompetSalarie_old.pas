{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 11/02/2004
Modifié le ... :   /  /
Description .. : Source TOF des fiches de saisie de compétences par : salarié,intérimaire,candidat
Mots clefs ... : TOF;PGSAISIECOMPETSALARIE
*****************************************************************}
Unit UTofPGSaisieCompetSalarie ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,
{$else}
     eMul,uTob,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,P5Def,EntPaie,SaisieList,uTableFiltre,PGOutilsFormation ;

Type
  TOF_PGSAISIECOMPETSALARIE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    NatureSaisie : String;
    TF : TTableFiltre;
    procedure OnClickSalarieSortie(Sender: TObject);
    procedure AffichageNomCaption(Sender : TObject);
  end ;

Type
  TOF_PGSAISIECOMPETSTAGE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    NatureSaisie : String;
    TF : TTableFiltre;
    procedure AffichageNomCaption(Sender : TObject); 
  end ;

Implementation

procedure TOF_PGSAISIECOMPETSALARIE.OnLoad ;
var DateArret : TDateTime;
    StDateArret : String;
begin
  Inherited ;
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
end;
end ;

procedure TOF_PGSAISIECOMPETSALARIE.OnArgument (S : String ) ;
var Num : Integer;
    Check : TCheckBox;
begin
  Inherited ;
TF  :=  TFSaisieList(Ecran).LeFiltre;
NatureSaisie := ReadTokenPipe(S,';');
SetControlText('XX_WHERE','');
For Num := 1 to VH_Paie.PGNbreStatOrg do
begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSI_TRAVAILN'+IntToStr(Num)),GetControl ('TPSI_TRAVAILN'+IntToStr(Num)));
end;
VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
If NatureSaisie = 'SALARIES' then TFSaisieList(Ecran).Caption := 'Saisie des compétences par salariés';
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
UpdateCaption(TFSaisieList(Ecran));
Check := TCheckBox(GetControl('CPRESENT'));
If Check <> Nil then Check.OnClick:=OnClickSalarieSortie;
SetControlEnabled('DATEPRESENCE',False);
SetControlText('DATEPRESENCE',DateToStr(Date));
TF.OnSetNavigate  :=  AffichageNomCaption;
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
        If TF.GetValue('PSI_TYPEINTERIM') = 'INT' then
        TFSaisieList(Ecran).Caption := 'Saisie des compétences pour l''intérimaire : '+StNom
        else TFSaisieList(Ecran).Caption := 'Saisie des compétences pour le stagiaire : '+StNom;
end;
UpdateCaption(TFSaisieList(Ecran));
end;


//TOF_PGSAISIECOMPETSTAGE

procedure TOF_PGSAISIECOMPETSTAGE.OnLoad ;
begin
  Inherited ;

end ;

procedure TOF_PGSAISIECOMPETSTAGE.OnArgument (S : String ) ;
var Num : Integer;
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


Initialization
  registerclasses ( [ TOF_PGSAISIECOMPETSALARIE,TOF_PGSAISIECOMPETSTAGE ] ) ;
end.

