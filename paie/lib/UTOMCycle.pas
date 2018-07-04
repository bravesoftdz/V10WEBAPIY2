{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Source TOF des FICHES : CYCLE_FSL et CYCLE
Mots clefs ... : TOF;CYCLE_FSL;CYCLE
*****************************************************************
PT1  17/07/2007  FLO  Gestion de la duplication d'une entête
PT2  09/08/2007  FLO  Recalcul automatique des compteurs lors d'une modification/suppression
}
unit UTOMCycle;

interface
uses Classes, SySUtils, Controls, UTOM, UTOF,Hctrls, HEnt1, uTableFiltre, SaisieList,
      {$IFDEF EAGLCLIENT}
      MaineAgl,EFiche,
      {$ELSE}
      FE_Main,Fiche, HDB, Db, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$ENDIF}
      HTB97;

Type
  TOM_CYCLEENTETE = Class(TOM)
  procedure OnArgument(stArgument: string);  override;
  procedure OnLoadRecord ;                   override; //PT1
  procedure OnAfterUpdateRecord;             override;
  procedure OnAfterDeleteRecord;             override; //PT2
  procedure OnDeleteRecord;                  override;
  procedure OnNewRecord;                     override; //PT1
  procedure OnUpdateRecord;                  override; //PT1
  private
    Data        : String;    //PT1
    Cycle       : String;    //PT2
    DateModif   : TDateTime; //PT2
    VerifImpact : Boolean;   //PT2
    procedure ClickPgColor (Sender : TObject);   //PT1
    procedure ClickPgDefaire (Sender : TObject); //PT1
  End;

Type
  TOM_CYCLE = Class(TOM)
  procedure OnArgument(stArgument: string);  override;
  procedure OnNewRecord;                     override;
  procedure OnUpdateRecord;                  override;
  procedure OnDeleteRecord;                  override; //PT1
  procedure OnAfterUpdateRecord;             override; //PT2
  procedure OnAfterDeleteRecord;             override; //PT2
  Private
    TF        : TTableFiltre;
    Cycle     : String;    //PT2
    DateModif : TDateTime; //PT2
  End;

type
  TOF_PGCYCLEENTETE = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad                   ;     override ;
    procedure DoClick(Sender : TObject);
    procedure DoDblClick(Sender : TObject);
    procedure Titreecran(Sender : TObject);
  private
    TF : TTableFiltre;
  end;

implementation

uses
 PGPlanningPresenceCycle, PgPlanningOutils, Graphics, Dialogs, UTOB, HMSgBox, PgOutils2, PGPresence;

{ TOM_CYCLEENTETE }

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Ouverture de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOM_CYCLEENTETE.OnArgument(stArgument: string);
Var Action : String;
    Btn    : TToolBarButton97;
begin
  inherited;

  Btn := TToolBarButton97(GetControl('BTNPGCOLORCYCLE'));
  if Assigned(Btn) then Btn.OnClick := ClickPgColor;

  Btn := TToolBarButton97(GetControl('BTNPGDEFAIRECOLORCYCLE'));
  if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;


  ReadTokenSt(stArgument);
  ReadTokenSt(stArgument);
  Action := ReadTokenSt(stArgument); // Le mode duplication est le 3e argument
  If Action = 'DUPLICATION' Then
  Begin
     SetControlVisible('BInsert',    False);
     SetControlVisible('BDelete',    False);
     Data := stArgument;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Chargement des données
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLEENTETE.OnLoadRecord ;
var
  EditColor : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
  Q : TQuery;
begin
  Inherited ;

  // Cas de la duplication
  If Data <> '' Then
  Begin
     Q := OpenSQL('SELECT * FROM CYCLEENTETE WHERE PYC_CYCLE="'+ReadTokenSt(Data)+'" AND PYC_DATEVALIDITE="'+UsDateTime(StrToDate(ReadTokenSt(Data)))+'"', True);
     If Not Q.EOF Then
     Begin
          SetField('PYC_DATEVALIDITE',  Date);
          SetField('PYC_LIBELLE',       Q.FindField('PYC_LIBELLE').AsString);
          SetField('PYC_TYPECYCLE',     Q.FindField('PYC_TYPECYCLE').AsString);
          SetField('PYC_ETATCYCLE',     Q.FindField('PYC_ETATCYCLE').AsString);
          SetField('PYC_NBCYCLE',       Q.FindField('PYC_NBCYCLE').AsInteger);
          SetField('PYC_COLORCYCLE',    Q.FindField('PYC_COLORCYCLE').AsString);
     End;
     Ferme(Q);
  End;

  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PYC_COLORCYCLE'));
  {$ELSE}
  Edit := THDBEdit(GetControl('PYC_COLORCYCLE'));
  {$ENDIF}
  If Assigned(Edit) then
    Begin
    if Edit.text<>'' then
      Begin
      Edit.Font.Color := StringToColor(Edit.Text);
      EditColor := THEdit(GetControl('PGCOLORCYCLE'));
      if Assigned(EditColor) then
        Begin
        EditColor.Color := StringToColor(Edit.Text);
        EditColor.Enabled := False;
        End;
      End;
    End;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Création d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLEENTETE.OnNewRecord;
begin
  inherited;
     SetField('PYC_COLORCYCLE','$00804000');
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Création ou modification d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLEENTETE.OnUpdateRecord;
begin
  inherited;
     //PT2 - Début
     If (LastError = 0) Then
     Begin
          VerifImpact := (ds.State = dsEdit);
          Cycle       := GetField('PYC_CYCLE');
          DateModif   := GetField('PYC_DATEVALIDITE');
     End;
     //PT2 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 03/07/2007
Modifié le ... : 09/08/2007
Description .. : Après mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLEENTETE.OnAfterUpdateRecord;
var DD,DF : TDatetime;
begin
  inherited;
     TFFiche(Ecran).Retour := GetField('PYC_CYCLE');
     ChargementTablette(TFFiche(Ecran).TableName, '');
     //PT2 - Début
     If (LastError = 0) And VerifImpact And (PresenceCycleIsUsed('CYC', Cycle)) Then
     begin
         PresenceDonneMoisCalculActuel (DD,DF);
         CompteursARecalculer(DD);
     end;
     //PT2 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 03/07/2007
Modifié le ... :   /  /
Description .. : Suppression
Mots clefs ... :
*****************************************************************}
procedure TOM_CYCLEENTETE.OnDeleteRecord;
begin
  inherited;

     If ExisteSQL('SELECT 1 FROM PROFILPRESENCE WHERE PPQ_TYPEAFFECT = "CYC" AND PPQ_CYCLEAFFECT ="'+GetField('PYC_CYCLE')+'"') Then
     Begin
          PGIBox (TraduireMemoire('Ce rythme de travail est affecté à un profil de présence.'), TraduireMemoire('Suppression impossible'));
          LastError := 1;
          Exit;
     End;

     ExecuteSQL('DELETE FROM CYCLE WHERE PYD_CYCLE="'+GetField('PYC_CYCLE')+'"');
     TFFiche(Ecran).Retour := GetField('PYC_CYCLE');

     //PT2 - Début
     If (LastError = 0) Then
     Begin
          Cycle     := GetField('PYC_CYCLE');
          DateModif := GetField('PYC_DATEVALIDITE');
     End;
     //PT2 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrment
Mots clefs ... :
*****************************************************************}
procedure TOM_CYCLEENTETE.OnAfterDeleteRecord;
var DD,DF : TDatetime;
begin
  inherited;
     If (LastError = 0) And (PresenceCycleIsUsed('CYC', Cycle)) Then
     begin
        PresenceDonneMoisCalculActuel (DD,DF);
        CompteursARecalculer(DD);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007 / PT1
Modifié le ... :   /  /    
Description .. : Choix d'une couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLEENTETE.ClickPgColor(Sender: TObject);
Var
  Col : String;
  Colors : TColorDialog;
  EditCol : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
     Colors:=TColorDialog.Create(Ecran);
     if Colors=nil then exit;

     if (Colors.Execute) then
     Begin
          Col:=ColorToString(Colors.Color);
          EditCol := THEdit(GetControl('PGCOLORCYCLE'));
          {$IFDEF EAGLCLIENT}
               Edit := THEdit(GetControl('PYC_COLORCYCLE'));
          {$ELSE}
               Edit := THDBEdit(GetControl('PYC_COLORCYCLE'));
          {$ENDIF}
          If Assigned(Edit) then
          Begin
               Edit.Font.Color := Colors.Color;
               if DS.State in [DsBrowse] then Ds.Edit;
               if GetField('PYC_COLORCYCLE') <> Col then
               Begin
                    SetField('PYC_COLORCYCLE',Col);
                    if Assigned(EditCol) then EditCol.Color := Colors.Color;
               End;
          End;
     End;
     Colors.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007
Modifié le ... :   /  /    
Description .. : Rétablissement de la couleur par défaut
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLEENTETE.ClickPgDefaire(Sender: TObject);
Var
  stVal : String;
  EditCol : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
     stVal := '$00804000';
     EditCol := THEdit(GetControl('PGCOLORCYCLE'));
     {$IFDEF EAGLCLIENT}
          Edit := THEdit(GetControl('PYC_COLORCYCLE'));
     {$ELSE}
          Edit := THDBEdit(GetControl('PYC_COLORCYCLE'));
     {$ENDIF}
     If Assigned(Edit) then
     Begin
          Edit.Font.Color := StringToColor(StVal);
          if DS.State in [DsBrowse] then Ds.Edit;
          if GetField('PYC_COLORCYCLE') <> StVal then
          Begin
               SetField('PYC_COLORCYCLE',StVal);
               if Assigned(EditCol) then EditCol.Color := StringToColor(StVal);
          End;
     End;
end;

{ TOM_CYCLE }

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /
Description .. : Chargement de l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLE.OnArgument(stArgument: string);
begin
  inherited;
  If (Ecran <> nil ) and (Ecran is TFSaisieList ) then
     TF := TFSaisieList(Ecran).LeFiltre;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /
Description .. : Nouvel enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_CYCLE.OnNewRecord;
begin
  inherited;
  if TF<>nil then
    Begin
    SetField('PYD_CYCLE',TF.TOBFiltre.GetValue('PYC_CYCLE'));
    SetField('PYD_ORDRECYCLE',TFSaisieList(Ecran).FListe.RowCount-1);
    //TFSaisieList(Ecran).FListe.ColEditables[0] := False;
    //TFSaisieList(Ecran).FListe.ColEditables[1] := False;
    SetField('PYD_NBMODCYCLE', 1);
    setfocuscontrol('PYD_MODELECYCLE');
    SetControlEnabled('PYD_CYCLE',False);
    SetControlEnabled('PYD_ORDRECYCLE',False);
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. : Création ou mise à jour d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLE.OnUpdateRecord;
begin
  inherited;

    if  GetField('PYD_MODELECYCLE') = '' Then
    Begin
        PGIError(TraduireMemoire('Le modèle de cycle doit être renseigné.'));
        LastError := 1;
    End;
     
    
    If LastError = 0 then
    begin

      //PT1 - Début
      If TF <> Nil Then
      Begin
          // Mise à jour du nombre de jours du rythme
          If Ds.State = dsInsert Then
          Begin
               ExecuteSQL('UPDATE CYCLEENTETE SET PYC_NBCYCLE=(PYC_NBCYCLE+1) WHERE PYC_CYCLE="'+TF.TOBFiltre.GetValue('PYC_CYCLE')+'" AND PYC_DATEVALIDITE="'+UsDateTime(TF.TOBFiltre.GetValue('PYC_DATEVALIDITE'))+'"');
               TF.RefreshEntete(TF.TOBFiltre.GetValue('PYC_CYCLE'));
          End;

          //PT2 - Début
          If (LastError = 0) Then
          Begin
               Cycle     := TF.TOBFiltre.GetValue('PYC_CYCLE');
               DateModif := TF.TOBFiltre.GetValue('PYC_DATEVALIDITE');
          End;
          //PT2 - Fin
      End;
    //PT1 - Fin
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Suppression d'un cycle
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLE.OnDeleteRecord;
begin
  inherited;
     // Seule la suppression du dernier élément est possible
     If GetField('PYD_ORDRECYCLE') <> (TFSaisieList(Ecran).FListe.RowCount-1) Then
     Begin
          PGIError(TraduireMemoire('Seul le dernier élément peut être supprimé.'));
          LastError := 1;
     End
     Else
     Begin
          // Mise à jour du nombre de cycles dans le rythme
          ExecuteSQL('UPDATE CYCLEENTETE SET PYC_NBCYCLE=(PYC_NBCYCLE-1) WHERE PYC_CYCLE="'+TF.TOBFiltre.GetValue('PYC_CYCLE')+'" AND PYC_DATEVALIDITE="'+UsDateTime(TF.TOBFiltre.GetValue('PYC_DATEVALIDITE'))+'"');
          TF.RefreshEntete(TF.TOBFiltre.GetValue('PYC_CYCLE'));
     End;

     //PT2 - Début
     If (TF <> Nil) And (LastError = 0) Then
     Begin
          Cycle     := TF.TOBFiltre.GetValue('PYC_CYCLE');
          DateModif := TF.TOBFiltre.GetValue('PYC_DATEVALIDITE');
     End;
     //PT2 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /    
Description .. : Suite à la création ou la mise à jour d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLE.OnAfterUpdateRecord;
var DD,DF: TDatetime;
begin
  inherited;
    If (LastError = 0) And (PresenceCycleIsUsed('CYC', Cycle)) Then
    begin
        PresenceDonneMoisCalculActuel (DD,DF);
        CompteursARecalculer(DD);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_CYCLE.OnAfterDeleteRecord;
var DD,DF : Tdatetime;
begin
  inherited;
    If (LastError = 0) And (PresenceCycleIsUsed('CYC', Cycle)) Then
    begin
      PresenceDonneMoisCalculActuel (DD,DF);   //PT3
      CompteursARecalculer(DD);
    end;
end;

{ TOF_PGCYCLEENTETE }

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Chargement de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCYCLEENTETE.OnArgument(Arguments: string);
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

procedure TOF_PGCYCLEENTETE.OnLoad;
begin
  inherited;
     TFSaisieList(Ecran).BCherche.Click;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Adaptation du titre
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGCYCLEENTETE.Titreecran(sender : Tobject);
var
stnom : String;

begin
  // affichage libellé
  StNom := TF.TobFiltre.GetValue('PYC_CYCLE') + ' ' + TF.TobFiltre.GetValue('PYC_LIBELLE');
  TFSaisieList(Ecran).Caption := 'Rythme de travail : '+StNom;
  UpdateCaption(TFSaisieList(Ecran));
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Gestion des boutons
Mots clefs ... :
*****************************************************************}
procedure TOF_PGCYCLEENTETE.DoClick(Sender: TObject);
Var
  Cycle : String;
  T        : TOB;
begin
  If UpperCase(TControl(Sender).Name) = 'BNEWENT' then
  Begin
    Cycle := AglLanceFiche('PAY','CYCLEENTETE','','','ACTION=CREATION');
    if Cycle <> '' then TF.RefreshEntete(Cycle);
  End;

  If UpperCase(TControl(Sender).Name) = 'B_PLANNING' then
    // affichage du planning
    PGPlanningPresenceCycleOpen(Date,(Date+30),'','"'+ TF.TOBFiltre.GetValue('PYC_CYCLE') +'"', [piaDemiJournee, piaSemaine, piaMois, piaHeure, piaOutlook], 'CYC');  //

  //PT1 - Début
  // Duplication du rythme de travail
  If UpperCase(TControl(Sender).Name) = 'BDUPLIQUERENT' Then
  Begin
    Cycle := AglLanceFiche('PAY','CYCLEENTETE','','','ACTION=CREATION;MONOFICHE;DUPLICATION;'+TF.TOBFiltre.GetValue('PYC_CYCLE')+';'+DateToStr(TF.TOBFiltre.GetValue('PYC_DATEVALIDITE')));
    if Cycle <> '' Then
    Begin
      // Duplication des cycles associés
      Try
          T := TOB.Create('LesCycles', Nil, -1);
          T.LoadDetailDBFromSQL('CYCLE','SELECT "'+Cycle+'" AS PYD_CYCLE, PYD_ORDRECYCLE, PYD_MODELECYCLE,PYD_NBMODCYCLE FROM CYCLE WHERE PYD_CYCLE="'+TF.TOBFiltre.GetValue('PYC_CYCLE')+'"');
          T.SetAllModifie(True);
          T.InsertOrUpdateDB();
      Except
          PGIInfo(TraduireMemoire('Les cycles n''ont pas été repris.'));
      End;
      If Assigned(T) Then FreeAndNil(T);

      // Rafraîchissement de l'écran
      TF.RefreshEntete(Cycle);
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
procedure TOF_PGCYCLEENTETE.DoDblClick(Sender: TObject);
Var
  Cycle : String;
begin
If UpperCase(TControl(Sender).Name) = 'TREEENTETE' then
  Begin
  Cycle := AglLanceFiche('PAY','CYCLEENTETE','',TF.TOBFiltre.GetValue('PYC_CYCLE'),'ACTION=MODIFICATION');
  if Cycle <> '' then TF.RefreshEntete(Cycle);
  End;
end;

initialization
  registerclasses([TOM_CYCLEENTETE,TOM_CYCLE,TOF_PGCYCLEENTETE]);

end.
