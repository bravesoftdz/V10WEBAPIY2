{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 17/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SAISIEFRAISSAL ()
Mots clefs ... : TOF;SAISIEFRAISSAL
*****************************************************************
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPGSaisieFraisSal;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     sysutils,ComCtrls,HCtrls,HMsgBox,UTOF,uTableFiltre,
     SaisieList,UTOB,LookUp,PGOutilsFormation;


Type
  TOF_SAISIEFRAISSAL = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TF: TTableFiltre;
    MillesimeEC:String;
    TypeSaisie,LeStage,LaSession,LeMillesime:String;
    procedure TreeViewdblClick(Sender:TObject);
    procedure StageElipsisClick(Sender:TObject);
    end ;

Implementation


procedure TOF_SAISIEFRAISSAL.TreeViewdblClick(Sender:TObject);
var MyTreenode1:tTreeNode;
    Niveau:Integer;
    St:String;
begin
MyTreenode1:=TF.LeTreeView.Selected;
Niveau:=myTreeNode1.Level;
if Niveau=0 Then
   begin
   St:=TF.TOBFiltre.GetValue('PSS_CODESTAGE')+';'+TF.TOBFiltre.GetValue('PSS_MILLESIME');
   AglLanceFiche('PAY','STAGE','',St,'CONSULTATION');
   end;
if Niveau=1 Then
   begin
   St:=TF.TOBFiltre.GetValue('PSS_CODESTAGE')+';'+IntToStr(TF.TOBFiltre.GetValue('PSS_ORDRE'))+';'+TF.TOBFiltre.GetValue('PSS_MILLESIME');
   AglLanceFiche('PAY','SESSIONSTAGE','',St,'CONSULTATION');
   end;
end;

procedure TOF_SAISIEFRAISSAL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEFRAISSAL.OnLoad ;
begin
  Inherited ;
TFSaisieList(Ecran).BCherche.Click;
end ;

procedure TOF_SAISIEFRAISSAL.OnArgument (S : String ) ;
var Q:TQuery;
    dateDebut,DateFin:TDateTime;
    Edit:THEdit;
    Combo:THValComboBox;
begin
  Inherited ;
TypeSaisie:=Trim(ReadTokenPipe(S,';'));
LeStage:=Trim(ReadTokenPipe(S,';'));
LaSession:=Trim(ReadTokenPipe(S,';'));
LeMillesime:=Trim(ReadTokenPipe(S,';'));
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
    TF := TFSaisieList(Ecran).LeFiltre;
  TF.LeTreeView.OnDblClick:=TreeViewdblClick;
MillesimeEC := RendMillesimeRealise(DateDebut,DateFin);
SetControlText('PSS_DATEDEBUT',DateToStr(DateDebut));
SetControlText('PSS_DATEFIN',DateToStr(DateFin));
Edit:=THEdit(GetControl('PSS_CODESTAGE'));
If Edit<>Nil Then Edit.OnelipsisClick:=StageElipsisClick;
SetControltext('PSS_ORDRE',LaSession);
SetControlText('PSS_CODESTAGE',LeStage);
SetControltext('PSS_MILLESIME',LeMillesime);
If TypeSaisie='FRAIS' Then
   begin
   TF.WhereTable:='WHERE PFO_CODESTAGE=:PSS_CODESTAGE AND PFO_ORDRE=:PSS_ORDRE AND PFO_MILLESIME=:PSS_MILLESIME AND PFO_EFFECTUE="X"';
   end;
Q:=OpenSQL('SELECT PST_DUREESTAGE,PSS_LIBELLE,PST_LIBELLE,PST_LIBELLE1,PSS_DATEDEBUT,PSS_DATEFIN,PST_DUREESTAGE,PST_JOURSTAGE'+
           ',PST_NATUREFORM,PST_LIEUFORM FROM SESSIONSTAGE'+
           ' LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME'+
           ' WHERE PSS_CODESTAGE="'+LeStage+'" AND PSS_MILLESIME="'+LeMillesime+'" AND PSS_ORDRE='+LaSession,True);  //DB2
If not Q.eof Then
   begin
   SetControlText('STAGE',Q.FindField('PST_LIBELLE').AsString); 
   SetControlText('SESSION',Q.FindField('PSS_LIBELLE').AsString);
   SetControlText('DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsString);
   SetControlText('DATEFIN',Q.FindField('PSS_DATEFIN').AsString);
   Combo:=THValComboBox(GetControl('NATUREFORM'));
   If Combo<>Nil then Combo.Value:=Q.FindField('PST_NATUREFORM').AsString;
   Combo:=THValComboBox(GetControl('LIEUFORMATION'));
   If Combo<>Nil then Combo.Value:=Q.FindField('PST_LIEUFORM').AsString;
   end;
Ferme(Q);
end ;

procedure TOF_SAISIEFRAISSAL.StageElipsisClick(Sender:TObject);
var  StWhere,StOrder : String;               
begin
StWhere := 'PST_ACTIF="X" AND'+
    ' (PST_MILLESIME="'+MillesimeEC+'") OR (PST_MILLESIME="0000" AND '+
    'PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'"))';
StOrder := 'PST_MILLESIME,PST_LIBELLE';
LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,StOrder, True,-1);
end;

Initialization
  registerclasses ( [ TOF_SAISIEFRAISSAL ] ) ;
end.
