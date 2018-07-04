{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PROFILPRESSALARIEGRP ()
Mots clefs ... : TOF;PROFILPRESSALARIEGRP
*****************************************************************
PT1  10/08/2007  FLO  Recalcul automatique des compteurs lors d'une affectation
PT2  20/08/2007  FLO  Si idem pop, affectation du profil paramétré
}
Unit UTofPG_PROFILPRESSALARIEGRP;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,htb97 ; 

Type
  TOF_PROFILPRESSALARIEGRP = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    ListeSalaries, Action, Arg : String;
    IsErrorOccur : Boolean;
    Control_TYPPROFILPRES : THValCombobox;
    Control_DATEVALIDITE : THEdit;
    Control_PROFILPRES : THValCombobox;
    procedure OnClickValid (Sender : TObject);
    procedure OnChangeTYPPROFIL(Sender: TObject);
    procedure OnExitDATEVALIDITE(Sender: TObject);
  end ;

Implementation

uses
  ed_tools, PGPresence, PGPopulOutils{$IFNDEF EAGLCLIENT}, UTOB{$ENDIF};

procedure TOF_PROFILPRESSALARIEGRP.OnUpdate ;
var
  CurrentSalarie, StErr, Profil : String;
  GPres : TGestionPresence;
  TobProfilPres : TOB;
begin
  Inherited ;

     //PT2 - Début
     If GetControlText('PPZ_TYPPROFILPRES') = 'POP' Then
     Begin
          If Not CanUsePopulation (TYP_POPUL_PRES) Then
          Begin
               LastError := 1;
               PGIInfo(TraduireMemoire('Le type de population "PRE" n''est pas valide.'), Ecran.caption);
               Exit;
          End;

          // Création de l'objet contenant tous les profils dépendant des populations
          GPres := TGestionPresence.Create (True,True,False,False,False,0,0,True);
     End;
     //PT2 - Fin

  if Action = 'CREATION-MULTISELECTION' then
  begin
    IsErrorOccur := False;
    BEGINTRANS;
    while ListeSalaries <> '' do
    begin
      CurrentSalarie := Trim(ReadTokenPipe(ListeSalaries, '|'));
      while CurrentSalarie <> '' do
      begin
        try
          If GetControlText('PPZ_TYPPROFILPRES') = 'PER' Then Profil := GetControlText('PPZ_PROFILPRES')
          Else Profil := GPres.GetProfil(CurrentSalarie, AGLStrToDate(GetControlText('PPZ_DATEVALIDITE')), 'POP'); //PT2

          If Profil <> '' Then  //PT2
          Begin
               TobProfilPres := TOB.Create('PROFILPRESSALARIE', Nil, -1);
               TobProfilPres.PutValue('PPZ_SALARIE',        CurrentSalarie);
               TobProfilPres.PutValue('PPZ_DATEVALIDITE',   AGLStrToDate(GetControlText('PPZ_DATEVALIDITE')));
               TobProfilPres.PutValue('PPZ_TYPPROFILPRES',  GetControlText('PPZ_TYPPROFILPRES'));
               TobProfilPres.PutValue('PPZ_PROFILPRES',     Profil);
               TobProfilPres.InsertOrUpdateDB();
               FreeAndNil(TobProfilPres);
          End
          Else //PT2
          Begin
               IsErrorOccur := True;
               StErr := TraduireMemoire('Impossible de déterminer le profil de présence associé');
              (GetControl('LISTEANOMALIES') as THListBox).Items.Add(TraduireMemoire('Salarié ')+CurrentSalarie+' : '+ StErr);
          End;

          CompteursARecalculer(AGLStrToDate(GetControlText('PPZ_DATEVALIDITE')), CurrentSalarie) // PT1
        except
          On E: Exception do
          begin
            IsErrorOccur := True;
            if (GetControl('LISTEANOMALIES') <> nil) and (GetControl('LISTEANOMALIES') is THListBox) then
            begin
              if pos('duplicate', E.message) > 0 then
                StErr := TraduireMemoire('Un profil est déjà associé à ce salarié pour cette date.')
              else
                StErr := E.message;
              (GetControl('LISTEANOMALIES') as THListBox).Items.Add(TraduireMemoire('Salarié ')+CurrentSalarie+' : '+ StErr);
            end;
          end;
        end;
        CurrentSalarie := Trim(ReadTokenPipe(ListeSalaries, '|'));
      end;
    end;
    if IsErrorOccur then
    begin
      (GetControl('PAGES') as THPageControl2).ActivePageIndex := 1;
      if PGIAsk(TraduireMemoire('Des erreurs sont survenues lors de la mise à jour.#10#13Voulez-vous enregistrer les modifications qui ont pu être effectuées ?'),Ecran.Caption) = mrYes then
        COMMITTRANS
      else
        ROLLBACK;
    end else COMMITTRANS;
  end;
  If Assigned (GPres) Then FreeAndNil(GPres); //PT2
  if not IsErrorOccur then Ecran.close else IsErrorOccur := False;
end ;

procedure TOF_PROFILPRESSALARIEGRP.OnLoad ;
begin
  Inherited ;
  if Action = 'CREATION-MULTISELECTION' then
  begin
    ListeSalaries := Trim(ReadTokenPipe(Arg, ';'));
    if pos('PSA_SALARIE',ListeSalaries) > 0 then
    begin
      ListeSalaries := StringReplace(ListeSalaries,'PSA_SALARIE','',[rfReplaceAll,rfIgnoreCase]);
      ListeSalaries := StringReplace(ListeSalaries,'OR','',[rfReplaceAll,rfIgnoreCase]);
      ListeSalaries := StringReplace(ListeSalaries,'=','',[rfReplaceAll,rfIgnoreCase]);
      ListeSalaries := StringReplace(ListeSalaries,'"','|',[rfReplaceAll,rfIgnoreCase]);
      ListeSalaries := StringReplace(ListeSalaries,' ','',[rfReplaceAll,rfIgnoreCase]);
      ListeSalaries := StringReplace(ListeSalaries,'||','|',[rfReplaceAll,rfIgnoreCase]);
    end;

  end;
  OnChangeTYPPROFIL(Self);
  OnExitDATEVALIDITE(Self);
end ;

procedure TOF_PROFILPRESSALARIEGRP.OnArgument (S : String ) ;
var
  stTemp : String;
begin
  Inherited ;
  Arg := S;
  stTemp := Trim(ReadTokenPipe(Arg, ';')); //On récupère le type d'action
  Action := Trim(ReadTokenPipe(stTemp, '='));
  Action := Trim(stTemp);
  (GetControl('BValider') as TToolBarButton97).OnClick := OnClickValid;
  Control_TYPPROFILPRES := (GetControl('PPZ_TYPPROFILPRES') as THValCombobox);
  Control_DATEVALIDITE := (GetControl('PPZ_DATEVALIDITE') as THEdit);
  Control_PROFILPRES := (GetControl('PPZ_PROFILPRES') as THValCombobox);
  Control_TYPPROFILPRES.OnChange := OnChangeTYPPROFIL;
  Control_DATEVALIDITE.OnExit := OnExitDATEVALIDITE;
end ;

procedure TOF_PROFILPRESSALARIEGRP.OnClickValid(Sender: TObject);
begin
  Self.OnUpdate;
end;

procedure TOF_PROFILPRESSALARIEGRP.OnChangeTYPPROFIL(Sender: TObject);
begin
  if Control_TYPPROFILPRES.Value = 'PER' then
  begin
    SetControlEnabled('PPZ_PROFILPRES',True);
//    SetField('PPZ_PROFILPRES', GetProfilIdem()); On doit initialiser le champs avec le profil qui correspond au Idem choisi.
  end else begin
    SetControlEnabled('PPZ_PROFILPRES',False);
  end;

end;

procedure TOF_PROFILPRESSALARIEGRP.OnExitDATEVALIDITE(Sender: TObject);
begin
  Control_PROFILPRES.plus := ' PPQ_DATEVALIDITE <= "'+USDateTime(AGLStrToDate(getControlText('PPZ_DATEVALIDITE')))+'" ';
end;

Initialization
  registerclasses ( [ TOF_PROFILPRESSALARIEGRP ] ) ; 
end.
