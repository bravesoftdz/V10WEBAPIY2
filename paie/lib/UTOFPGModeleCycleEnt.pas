{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MODELECYCLE_FSL ()
Mots clefs ... : TOF;MODELECYCLE_FSL
*****************************************************************
PT1  17/07/2007  FLO  Ajout d'une duplication d'un modèle complet
}
unit UTOFPGModeleCycleEnt;

interface

uses
  SysUtils,Classes,Controls,
  Utof, uTableFiltre,Hctrls, SaisieList,HTB97,DB
  {$IFDEF EAGLCLIENT}
  ,MaineAgl
  {$ELSE}
  ,FE_Main
  {$ENDIF};

type
  TOF_PGMODELECYCLEENT = class(TOF)
    procedure OnArgument(Arguments: string)  ; override ;
    procedure OnLoad                         ; override ;
  private
    TF : TTableFiltre;
    procedure DoClick(Sender : TObject);
    procedure DoDblClick(Sender : TObject);
    procedure Titreecran(Sender : TObject);
  end;


implementation

uses
 PGPlanningPresenceCycle, PgPlanningOutils, HMsgBox, HEnt1, UTOB;

{ TOF_PGMODELECYCLEENT }

procedure TOF_PGMODELECYCLEENT.OnArgument(Arguments: string);
begin
  inherited;
  If (Ecran <> nil ) and (Ecran is TFSaisieList ) then
     TF := TFSaisieList(Ecran).LeFiltre;

  TF.OnSetNavigate  := Titreecran;

  TFSaisieList(Ecran).TreeEntete.OnDblClick := DoDblClick;

  TToolBarButton97(GetControl('BNEWENT')).OnClick := DoClick;
  TToolBarButton97(GetControl('B_PLANNING')).OnClick := DoClick;
  TToolBarButton97(GetControl('BDUPLIQUERENT')).OnClick := DoClick; //PT1
end;


{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Gestion des boutons
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMODELECYCLEENT.DoClick(Sender: TObject);
Var
  ModCycle : String;
  T        : TOB;
begin
  If UpperCase(TControl(Sender).Name) = 'BNEWENT' then
  Begin
    ModCycle := AglLanceFiche('PAY','MODELECYCLE','','','ACTION=CREATION;MONOFICHE');
    if ModCycle <> '' then TF.RefreshEntete(ModCycle);
  End;

  If UpperCase(TControl(Sender).Name) = 'B_PLANNING' then
    // affichage du planning
    PGPlanningPresenceCycleOpen(Date,(Date+30),'','"'+ TF.TOBFiltre.GetValue('PMY_MODELECYCLE') +'"', [piaDemiJournee, piaSemaine, piaMois, piaHeure, piaOutlook], 'MOD');  //

  //PT1 - Début
  // Duplication de modèle de cycle
  If UpperCase(TControl(Sender).Name) = 'BDUPLIQUERENT' Then
  Begin
    ModCycle := AglLanceFiche('PAY','MODELECYCLE','','','ACTION=CREATION;MONOFICHE;DUPLICATION;'+TF.TOBFiltre.GetValue('PMY_MODELECYCLE')+';'+DateToStr(TF.TOBFiltre.GetValue('PMY_DATEVALIDITE')));
    if ModCycle <> '' Then
    Begin
      // Duplication des cycles
      Try
          T := TOB.Create('LesModeles', Nil, -1);
          T.LoadDetailDBFromSQL('MODELECYCLE','SELECT "'+ModCycle+'" AS PMO_MODELECYCLE, PMO_JOURNEETYPE, PMO_ORDREJOUR FROM MODELECYCLE WHERE PMO_MODELECYCLE="'+TF.TOBFiltre.GetValue('PMY_MODELECYCLE')+'"');
          T.SetAllModifie(True);          
          T.InsertOrUpdateDB();
      Except
          PGIInfo(TraduireMemoire('Les modèles de cycle n''ont pas été repris.'));
      End;
      If Assigned(T) Then FreeAndNil(T);

      // Rafraîchissement de l'écran
      TF.RefreshEntete(ModCycle);
    End;
  End;
  //PT1 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Gestion du double-clic dans l'arborescence
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGMODELECYCLEENT.DoDblClick(Sender: TObject);
Var
  ModCycle : String;
begin
If UpperCase(TControl(Sender).Name) = 'TREEENTETE' then
  Begin
  ModCycle := AglLanceFiche('PAY','MODELECYCLE','',TF.TOBFiltre.GetValue('PMY_MODELECYCLE'),'ACTION=MODIFICATION;MONOFICHE');
  if ModCycle <> '' then TF.RefreshEntete(ModCycle);
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Adaptation du titre
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGMODELECYCLEENT.TitreEcran(sender : Tobject);
var
stnom : String;

begin
  // affichage libellé
  StNom := TF.TobFiltre.GetValue('PMY_MODELECYCLE') + ' ' + TF.TobFiltre.GetValue('PMY_LIBELLE');
  TFSaisieList(Ecran).Caption := 'Modèle de cycle : '+StNom;
  UpdateCaption(TFSaisieList(Ecran));
end;


procedure TOF_PGMODELECYCLEENT.OnLoad;
begin
  inherited;
   TFSaisieList(Ecran).BCherche.Click;
end;

initialization
  registerclasses([TOF_PGMODELECYCLEENT]);
end.
