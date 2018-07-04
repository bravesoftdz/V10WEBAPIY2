{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Source TOF des FICHES : MODELECYCLEENT et MODELECYCLE ()
Mots clefs ... : TOF;MODELECYCLEENT;MODELECYCLE
*****************************************************************
PT1  17/07/2007  FLO  Gestion de la duplication d'une entête
PT2  09/08/2007  FLO  Recalcul automatique des compteurs lors d'une modification/suppression
PT3  27/08/2007  FLO  suppression impossible si modèle utilisé
}
unit UTOMModeleCycleEnt;

interface

uses
   SysUtils,StdCtrls,Classes,  Dialogs, Graphics,HCtrls,HMSgBox,uTableFiltre,SaisieList,
   UTOM,UTOB,HEnt1, HTB97
   {$IFDEF EAGLCLIENT}
   ,EFiche
   {$ELSE}
   ,Fiche, HDB, Db {$IFNDEF DBXPRESS}, dbtables {$ELSE}, uDbxDataSet {$ENDIF}
   {$ENDIF}
   ;

Type
  TOM_MODELECYCLEENT = Class(TOM)
  procedure OnArgument(stArgument: string); override;
  procedure OnAfterUpdateRecord;            override;
  procedure OnAfterDeleteRecord;            override; //PT2
  procedure OnLoadRecord ;                  override;
  procedure OnDeleteRecord;                 override;
  procedure OnNewRecord;                    override; //PT1
  procedure OnUpdateRecord;                 override; //PT1
  Private
    Data        : String;
    Cycle       : String;    //PT2
    DateModif   : TDateTime; //PT2
    VerifImpact : Boolean;   //PT2
    procedure ClickPgColor (Sender : TObject);
    procedure ClickPgDefaire (Sender : TObject);
  End;

Type
  TOM_MODELECYCLE = Class(TOM)
  procedure OnArgument(stArgument: string); override;
  procedure OnLoadRecord ;                  override ;
  procedure OnNewRecord;                    override;
  procedure OnUpdateRecord;                 override;
  procedure OnDeleteRecord;                 override; //PT1
  procedure OnAfterUpdateRecord;            override; //PT2
  procedure OnAfterDeleteRecord;            override; //PT2
  Private
   TF        : TTableFiltre;
   Cycle     : String; //PT2
   DateModif : TDateTime; //PT2
  Function PgFloatToStrTime( DH: Double; Format : String; b100: Boolean = False) : String;
  end ;

implementation
uses HeureUtil, PGPresence;


{ TOM_MODELECYCLEENT }
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 03/07/2007
Modifié le ... :   /  /    
Description .. : Chargement de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_MODELECYCLEENT.OnLoadRecord ;
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

  //PT1 - Début
  // Cas de la duplication
  If Data <> '' Then
  Begin
     Q := OpenSQL('SELECT * FROM MODELECYCLEENT WHERE PMY_MODELECYCLE="'+ReadTokenSt(Data)+'" AND PMY_DATEVALIDITE="'+UsDateTime(StrToDate(ReadTokenSt(Data)))+'"', True);
     If Not Q.EOF Then
     Begin
          SetField('PMY_DATEVALIDITE', Date);
          SetField('PMY_LIBELLE',       Q.FindField('PMY_LIBELLE').AsString);
          SetField('PMY_NBJOUR',        Q.FindField('PMY_NBJOUR').AsInteger);
          SetField('PMY_COLORMODCYCLE', Q.FindField('PMY_COLORMODCYCLE').AsString);
     End;
     Ferme(Q);
  End;
  //PT1 - Fin

  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PMY_COLORMODCYCLE'));
  {$ELSE}
  Edit := THDBEdit(GetControl('PMY_COLORMODCYCLE'));
  {$ENDIF}
  If Assigned(Edit) then
    Begin
    if Edit.text<>'' then
      Begin
      Edit.Font.Color := StringToColor(Edit.Text);
      EditColor := THEdit(GetControl('PGCOLORMODCYCLE'));
      if Assigned(EditColor) then
        Begin
        EditColor.Color := StringToColor(Edit.Text);
        EditColor.Enabled := False;
        End;
      End;
    End;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 03/07/2007
Modifié le ... :   /  /
Description .. : ON ARGUMENT
Mots clefs ... : 
*****************************************************************}
procedure TOM_MODELECYCLEENT.OnArgument(stArgument: string);
var
Btn  : TToolBarButton97;
Action : String;
begin
  inherited;

  Btn := TToolBarButton97(GetControl('BTNPGCOLORMODCYCLE'));
  if Assigned(Btn) then Btn.OnClick := ClickPgColor;

  Btn := TToolBarButton97(GetControl('BTNPGDEFAIRECOLORMODCYCLE'));
  if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;

  //PT1 - Début
  ReadTokenSt(stArgument);
  ReadTokenSt(stArgument);
  Action := ReadTokenSt(stArgument); // Le mode duplication est le 3e argument
  If Action = 'DUPLICATION' Then
  Begin
     SetControlVisible('BInsert',    False);
     SetControlVisible('BDelete',    False);
     Data := stArgument;
  End;
  //PT1 - Fin
end;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/07/2007 / PT1
Modifié le ... :   /  /    
Description .. : Création ou modification d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_MODELECYCLEENT.OnUpdateRecord;
begin
  inherited;
     //PT2 - Début
     If (LastError = 0) Then
     Begin
          VerifImpact := (ds.State = dsEdit);
          Cycle       := GetField('PMY_MODELECYCLE');
          DateModif   := GetField('PMY_DATEVALIDITE');
     End;
     //PT2 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 03/07/2007
Modifié le ... :   /  /    
Description .. : Après mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_MODELECYCLEENT.OnAfterUpdateRecord;
Var DD,DF : Tdatetime;
begin
  inherited;
     TFFiche(Ecran).Retour := GetField('PMY_MODELECYCLE');

     //PT2 - Début
     If (LastError = 0) And VerifImpact And PresenceCycleIsUsed('MOD', Cycle) Then
     begin
     PresenceDonneMoisCalculActuel(DD,DF);
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
procedure TOM_MODELECYCLEENT.OnDeleteRecord;
begin
  inherited;
     //PT3 - Début
     // Vérification que le modèle ne soit pas utilisé au sein d'un rythme de travail

     If ExisteSQL('SELECT 1 FROM CYCLE WHERE PYD_MODELECYCLE="'+GetField('PMY_MODELECYCLE')+'"') Then
     Begin
          PGIBox (TraduireMemoire('Le modèle est utilisé au sein d''un rythme de travail.'), TraduireMemoire('Suppression impossible'));
          LastError := 1;
          Exit;
     End;

      If ExisteSQL('SELECT 1 FROM PROFILPRESENCE WHERE PPQ_TYPEAFFECT = "MOD" AND PPQ_CYCLEAFFECT ="'+GetField('PMY_MODELECYCLE')+'"') Then
     Begin
          PGIBox (TraduireMemoire('Ce modèle est affecté à un profil de présence.'), TraduireMemoire('Suppression impossible'));
          LastError := 1;
          Exit;
     End;
     //PT3 - Fin

     ExecuteSQL('DELETE FROM MODELECYCLE WHERE PMO_MODELECYCLE="'+GetField('PMY_MODELECYCLE')+'"');
     TFFiche(Ecran).Retour := GetField('PMY_MODELECYCLE');

     //PT2 - Début
     If (LastError = 0) Then
     Begin
          Cycle     := GetField('PMY_MODELECYCLE');
          DateModif := GetField('PMY_DATEVALIDITE');
     End;
     //PT2 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_MODELECYCLEENT.OnAfterDeleteRecord;
var DD,DF : Tdatetime;
begin
  inherited;
     If (LastError = 0) And (PresenceCycleIsUsed('MOD', Cycle)) Then
     begin
       PresenceDonneMoisCalculActuel (DD,DF);
       CompteursARecalculer(DD);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Création d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_MODELECYCLEENT.OnNewRecord;
begin
  inherited;
     SetField('PMY_COLORMODCYCLE','$00804000');  
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 03/07/2007
Modifié le ... :   /  /    
Description .. : Clcick sur la couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_MODELECYCLEENT.ClickPgColor(Sender: TObject);
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
  EditCol := THEdit(GetControl('PGCOLORMODCYCLE'));
  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PMY_COLORMODCYCLE'));
  {$ELSE}
  Edit := THDBEdit(GetControl('PMY_COLORMODCYCLE'));
  {$ENDIF}
  If Assigned(Edit) then
    Begin
    Edit.Font.Color := Colors.Color;
    if DS.State in [DsBrowse] then Ds.Edit;
    if GetField('PMY_COLORMODCYCLE') <> Col then
       Begin
       SetField('PMY_COLORMODCYCLE',Col);
       if Assigned(EditCol) then EditCol.Color := Colors.Color;
       End;
    End;
  End;
Colors.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 03/07/2007
Modifié le ... :   /  /    
Description .. : Bouton supprimer choix de la couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_MODELECYCLEENT.ClickPgDefaire(Sender: TObject);
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
  EditCol := THEdit(GetControl('PGCOLORMODCYCLE'));
  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PMY_COLORMODCYCLE'));
  {$ELSE}
  Edit := THDBEdit(GetControl('PMY_COLORMODCYCLE'));
  {$ENDIF}
  If Assigned(Edit) then
    Begin
    Edit.Font.Color := StringToColor(StVal);
    if DS.State in [DsBrowse] then Ds.Edit;
    if GetField('PMY_COLORMODCYCLE') <> StVal then
      Begin
      SetField('PMY_COLORMODCYCLE',StVal);
      if Assigned(EditCol) then EditCol.Color := StringToColor(StVal);
      End;
    End;

end;

{ TOM_MODELECYCLE }

procedure TOM_MODELECYCLE.OnArgument(stArgument: string);
begin
  inherited;
  If (Ecran <> nil ) and (Ecran is TFSaisieList ) then
     TF := TFSaisieList(Ecran).LeFiltre;
end;

procedure TOM_MODELECYCLE.OnLoadRecord;
Var
 I : Integer;
 Tob_JourneeType,T1 : Tob;
 DMod,DPause, DpauseNoneff, Dnontraveff , Dureetraveff : Double;
 JourTyp : string;
begin
  inherited;
  if TF<>nil then
    Begin
    Tob_JourneeType := TOB.Create('JOURNEETYPE',nil,-1);
    Tob_JourneeType.LoadDetailFromSQL('SELECT * FROM JOURNEETYPE');
    DMod :=0; DPause :=0;  DpauseNoneff :=0; Dnontraveff := 0;

    For i := 0 to TF.TOBFiltre.Detail.Count -1 do
      Begin
      JourTyp := TF.TOBFiltre.Detail[i].GetValue('PMO_JOURNEETYPE');
      T1 := Tob_JourneeType.FindFirst(['PJO_JOURNEETYPE'],[JourTyp],False);
      if Assigned(T1) then
         Begin

         DMod   := DMod + T1.GetValue('PJO_DUREEPLAGE1') + T1.GetValue('PJO_DUREEPLAGE2');
         DPause := DPause +T1.GetValue('PJO_DUREEPAUSE');
         Dnontraveff := Dnontraveff + T1.GetValue('PJO_DUREENOTRAV');
         if T1.GetValue('PJO_PAUSEEFFECTIF') <> 'X' then
         DpauseNoneff := DpauseNoneff + T1.GetValue('PJO_DUREEPAUSE');
         End;
      End;
    If DMod<>0 then   SetControlText('DUREEMODELE',PgFloatToStrTime(DMod,'hh:nn',True));
    If DPause<>0 then SetControlText('DUREEPAUSE', PgFloatToStrTime(DPause,'hh:nn',True));
    Dureetraveff  := Dmod - DpauseNoneff - Dnontraveff;
    if dureetraveff <> 0 then SetcontrolText('DUREETRAVEFF', PgFloatToStrTime(dureetraveff, 'hh:mm', true));
    
    FreeAndNil(Tob_JourneeType);
    End;


end;

procedure TOM_MODELECYCLE.OnNewRecord;
begin
  inherited;
  if TF<>nil then
    Begin
    SetField('PMO_MODELECYCLE',TF.TOBFiltre.GetValue('PMY_MODELECYCLE'));
    SetField('PMO_ORDREJOUR',TFSaisieList(Ecran).FListe.RowCount-1);
    //TFSaisieList(Ecran).FListe.ColEditables[0] := False;
    //TFSaisieList(Ecran).FListe.ColEditables[1] := False;
    SetControlEnabled('PMO_MODELECYCLE',False);
    SetControlEnabled('PMO_ORDREJOUR',False);
    End;
end;

procedure TOM_MODELECYCLE.OnUpdateRecord;
begin
  inherited;
    if  GetField('PMO_JOURNEETYPE') = '' Then
    Begin
        PGIError(TraduireMemoire('La journée type doit être renseignée.'));
        LastError := 1;
    End;

    If LastError = 0 then
    begin

      if TF<>nil then
      Begin
          //PT1 - Début
          {
          if TFSaisieList(Ecran).FListe.RowCount-1 > TF.TOBFiltre.GetValue('PMY_NBJOUR') then
          Begin
               LastError := 1;
               PgiBox('Le cycle comprend '+IntToStr(TF.TOBFiltre.GetValue('PMY_NBJOUR'))+' jours. Création impossible',TFSaisieList(Ecran).caption);
          End;
          }

          // Mise à jour du nombre de jours du modèle de cycle
          If Ds.State = dsInsert Then
          Begin
               ExecuteSQL('UPDATE MODELECYCLEENT SET PMY_NBJOUR=(PMY_NBJOUR+1) WHERE PMY_MODELECYCLE="'+TF.TOBFiltre.GetValue('PMY_MODELECYCLE')+'" AND PMY_DATEVALIDITE="'+UsDateTime(TF.TOBFiltre.GetValue('PMY_DATEVALIDITE'))+'"');
               TF.RefreshEntete(TF.TOBFiltre.GetValue('PMY_MODELECYCLE'));
          End;
          //PT1 - Fin

          //PT2 - Début
          If (LastError = 0) Then
          Begin
               Cycle     := TF.TOBFiltre.GetValue('PMY_MODELECYCLE');
               DateModif := TF.TOBFiltre.GetValue('PMY_DATEVALIDITE');
          End;
          //PT2 - Fin
     End;
    End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 17/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Avant la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_MODELECYCLE.OnDeleteRecord;
begin
  inherited;
     // Seule la suppression du dernier élément est possible
     If GetField('PMO_ORDREJOUR') <> (TFSaisieList(Ecran).FListe.RowCount-1) Then
     Begin
          PGIError(TraduireMemoire('Seul le dernier élément peut être supprimé.'));
          LastError := 1;
     End
     Else
     Begin
          // Mise à jour du nombre de jours du modèle de cycle
          ExecuteSQL('UPDATE MODELECYCLEENT SET PMY_NBJOUR=(PMY_NBJOUR-1) WHERE PMY_MODELECYCLE="'+TF.TOBFiltre.GetValue('PMY_MODELECYCLE')+'" AND PMY_DATEVALIDITE="'+UsDateTime(TF.TOBFiltre.GetValue('PMY_DATEVALIDITE'))+'"');
          TF.RefreshEntete(TF.TOBFiltre.GetValue('PMY_MODELECYCLE'));
     End;

     //PT2 - Début
     If (TF <> Nil) And (LastError = 0) Then
     Begin
          Cycle     := TF.TOBFiltre.GetValue('PMY_MODELECYCLE');
          DateModif := TF.TOBFiltre.GetValue('PMY_DATEVALIDITE');
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
procedure TOM_MODELECYCLE.OnAfterUpdateRecord;
var DD,DF : Tdatetime;
begin
  inherited;
    If (LastError = 0) And (PresenceCycleIsUsed('MOD', Cycle)) Then
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
procedure TOM_MODELECYCLE.OnAfterDeleteRecord;
var DD,DF : TDateTime;
begin
  inherited;
    If (LastError = 0) And (PresenceCycleIsUsed('MOD', Cycle)) Then
    begin
      PresenceDonneMoisCalculActuel (DD,DF);
      CompteursARecalculer(DD);
    end;
end;

Function  TOM_MODELECYCLE.PgFloatToStrTime( DH: Double; Format : String; b100: Boolean = False) : String;
var
  DHi,DHa : Double;
Begin
 DHi := Int(DH);
 DHa := Frac(DH);
 If (DHa > 0) AND (DHi=0)  then
   Begin
   if Arrondi(DHa*60,0) < 10 then
     Result := '00:0' + FloatToStr(Arrondi(DHa*60,0))
   else
     Result := '00:' + FloatToStr(Arrondi(DHa*60,0));
   End
 else
   If (DHa > 0) AND (DHi>0)  then
   Begin
   IF DHi < 10 then
     Result := '0' + FloatToStr(DHi) + ':'
   else
     Result := FloatToStr(DHi) + ':' ;
   if Arrondi(DHa*60,0) < 10 then
     Result := Result + '0'+ FloatToStr(Arrondi(DHa*60,0))
   else
     Result := Result + FloatToStr(Arrondi(DHa*60,0));
   End
 else
   if (DHi>0) then
      Begin
      if DHi <10 then
      Result := '0'+ FloatToStr(DHi) + ':00'
      else
      Result := FloatToStr(DHi) + ':00'
      End
  else
      Result:='00:00';
end;

initialization
  registerclasses([TOM_MODELECYCLEENT,TOM_MODELECYCLE]);
end.
