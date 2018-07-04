{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/03/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PROFILPRESSALARIE (PROFILPRESSALARIE)
Mots clefs ... : TOM;PROFILPRESSALARIE
*****************************************************************
PT1  10/08/2007  FLO  Recalcul automatique des compteurs lors d'une affectation
PT2  20/08/2007  FLO  Si idem pop, affectation du profil paramétré
}
Unit UTomPGPROFILPRESSALARIE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, HDB,
     UTob ;

Type
  TOM_PROFILPRESSALARIE = Class (TOM)
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
  private
    Action, Arg : String;
    SalarieModif: String;    //PT1
    DateModif   : TDateTime; //PT1
{$IFNDEF EAGLCLIENT}
    Control_TYPPROFILPRES : THDBValComboBox;
    Control_DATEVALIDITE : THDBEdit;
    Control_PROFILPRES : THDBValCombobox;
{$else}
    Control_TYPPROFILPRES : THValCombobox;
    Control_DATEVALIDITE : THEdit;
    Control_PROFILPRES : THValCombobox;
{$ENDIF}
    procedure OnChangeTYPPROFIL (Sender : TObject);
    procedure OnExitDATEVALIDITE(Sender: TObject);
    procedure OnPlanningClick(Sender: TObject);
  end ;

Implementation
uses ed_tools, PGPlanningPresenceSal, PgPlanningOutils, HTB97, PGPresence, PGPopulOutils;

procedure TOM_PROFILPRESSALARIE.OnDeleteRecord ;
begin
  Inherited ;
  //PT1 - Début
  If (LastError=0) Then
  Begin
     SalarieModif := GetField('PPZ_SALARIE');
     DateModif    := GetField('PPZ_DATEVALIDITE');
  End;
  //PT1 - Fin
end ;

procedure TOM_PROFILPRESSALARIE.OnUpdateRecord ;
var Profil : String;
begin
  Inherited ;
     SalarieModif := GetField('PPZ_SALARIE');
     DateModif    := GetField('PPZ_DATEVALIDITE');
     
     If GetField('PPZ_TYPPROFILPRES') = 'PER'  // si profil personnalisé : controle saisie du  profil
     then
     begin
          if Getfield('PPZ_PROFILPRES') = '' then
          begin
               LastError := 1;
               PGIInfo(TraduireMemoire('Le profil de présence doit être renseigné.'), Ecran.caption);
               Setfocuscontrol('PPZ_PROFILPRES');
               Exit;
          end;

     end;
     
     //PT2 - Début
     If GetField('PPZ_TYPPROFILPRES') = 'POP' Then
     Begin
          If Not CanUsePopulation (TYP_POPUL_PRES) Then
          Begin
               LastError := 1;
               PGIInfo(TraduireMemoire('Le type de population "PRE" n''est pas valide.'), Ecran.caption);
               Exit;
          End;
     End;
     //PT2 - Fin

     //PT1 - Début
     If (LastError=0) Then
     Begin

          // PT2 - Début
          // Récupération du profil de la population si besoin
          If GetField('PPZ_TYPPROFILPRES') = 'POP' Then
          Begin
               Profil := GetProfilPresenceFromPop(SalarieModif, DateModif);
               If Profil <> '' Then
                    SetField('PPZ_PROFILPRES', Profil)
               Else
               Begin
                    LastErrorMsg := TraduireMemoire('Impossible de déterminer le profil de présence associé')+' '+TraduireMemoire(' à la population du salarié');
                    LastError := 1;
               End;
          End;
          // PT2 - Fin
     End;
     //PT1 - Fin


end ;

procedure TOM_PROFILPRESSALARIE.OnAfterUpdateRecord ;
begin
  Inherited ;
     If (LastError = 0) Then CompteursARecalculer(DateModif, SalarieModif); //PT1
end ;

procedure TOM_PROFILPRESSALARIE.OnAfterDeleteRecord ;
begin
  Inherited ;
     If (LastError = 0) Then CompteursARecalculer(DateModif, SalarieModif); //PT1
end ;

procedure TOM_PROFILPRESSALARIE.OnLoadRecord ;
begin
  Inherited ;
  if Action = 'CREATION' then
  begin
    SetControlText('PPZ_SALARIE',Trim(ReadTokenPipe(Arg, ';')));
  end;
  OnChangeTYPPROFIL(Self);
  OnExitDATEVALIDITE(Self);
end ;

procedure TOM_PROFILPRESSALARIE.OnArgument ( S: String ) ;
var
  stTemp : String;
begin
  Inherited ;
  Arg := S;
  stTemp := Trim(ReadTokenPipe(Arg, ';')); //On récupère le type d'action
  Action := Trim(ReadTokenPipe(stTemp, '='));
  Action := Trim(stTemp);
{$IFNDEF EAGLCLIENT}
  Control_TYPPROFILPRES := (GetControl('PPZ_TYPPROFILPRES') as THDBValComboBox);
  Control_DATEVALIDITE := (GetControl('PPZ_DATEVALIDITE') as THDBEdit);
  Control_PROFILPRES := (GetControl('PPZ_PROFILPRES') as THDBValCombobox);
{$else}
  Control_TYPPROFILPRES := (GetControl('PPZ_TYPPROFILPRES') as THValCombobox);
  Control_DATEVALIDITE := (GetControl('PPZ_DATEVALIDITE') as THEdit);
  Control_PROFILPRES := (GetControl('PPZ_PROFILPRES') as THValCombobox);
{$ENDIF}
  Control_TYPPROFILPRES.OnChange := OnChangeTYPPROFIL;
  Control_DATEVALIDITE.OnExit := OnExitDATEVALIDITE;
  if (GetControl('BPLANNING') is TToolbarButton97) then
    (GetControl('BPLANNING') as TToolbarButton97).OnClick := OnPlanningClick;

end ;

procedure TOM_PROFILPRESSALARIE.OnChangeTYPPROFIL(Sender: TObject);
begin
  if Control_TYPPROFILPRES.Value = 'PER' then
  begin
    SetControlEnabled('PPZ_PROFILPRES',True);
//    SetField('PPZ_PROFILPRES', GetProfilIdem()); On doit initialiser le champs avec le profil qui correspond au Idem choisi.
  end else begin
    SetControlEnabled('PPZ_PROFILPRES',False);
    //SetControlProperty('PPZ_PROFILPRES', 'Plus' ,' AND PPQ_DATEVALIDITE = (Select Max(2.PPQ_DATEVALIDITE) from PROFILPRESENCE 2 where 2.PPQ_PROFILPRES = PPQ_PROFILPRES and 2.PPQ_DATEVALIDITE <= "'+USDATETIME(AGLStrToDate(getControlText('PPZ_DATEVALIDITE')))+'")');
  end;

end;

procedure TOM_PROFILPRESSALARIE.OnExitDATEVALIDITE(Sender: TObject);
begin
  Control_PROFILPRES.plus := ' PPQ_DATEVALIDITE <= "'+USDateTime(getField('PPZ_DATEVALIDITE'))+'" ';
end;

procedure TOM_PROFILPRESSALARIE.OnPlanningClick(Sender: TObject);
var
  NiveauRupt: TNiveauRupture;
begin
  NiveauRupt.ChampsRupt[1] := '';
  NiveauRupt.ChampsRupt[2] := '';
  NiveauRupt.ChampsRupt[3] := '';
  NiveauRupt.ChampsRupt[4] := '';
  NiveauRupt.CondRupt := '';
  NiveauRupt.NiveauRupt := 0;

  PGPlanningPresenceSalOpen(Date,Date+30, '', '', NiveauRupt, [piaHeure, piaDemiJournee, piaSemaine, piaMois, piaOutlook], False, True, ' WHERE PSA_SALARIE = "'+getField('PPZ_SALARIE')+'"');   // NiveauRupt : TNiveauRupture;  ; DateReference : TDateTime = 0  ,StWhTobItems
end;

Initialization
  registerclasses ( [ TOM_PROFILPRESSALARIE ] ) ; 
end.
