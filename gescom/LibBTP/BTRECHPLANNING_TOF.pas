{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/03/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECHPLANNING ()
Mots clefs ... : TOF;BTRECHPLANNING
*****************************************************************}
Unit BTRECHPLANNING_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}
      dbtables,
     {$ELSE}
      uDbxDataSet,
     {$ENDIF}
     {$IFNDEF ERADIO}
      Fe_Main,
     {$ENDIF !ERADIO}
     mul,
{$ELSE}
     eMul,
     uTob,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF,
     Paramsoc,
     EntGC,
     FactComm,
     uEntCommun;




Type
  TOF_BTRECHPLANNING = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  //
  Affaire       : THEdit;
  Affaire0      : THEdit;
  Affaire1      : THEdit;
  Affaire2      : THEdit;
  Affaire3      : THEdit;
  Avenant       : THEdit;
  //
  Ressource     : THEdit;
  NumPhase      : THEdit;
  BTRecherche   : TToolBarButton97;
  //
  TAffaire      : THLabel;
  TRessource    : THLabel;
  TNumPhase     : THLabel;
  //
  ModePlanning  : String;
  TRech         : String;
  //
    procedure GetObjects;
    procedure SetScreenEvents;
    procedure Controlechamp(Champ, Valeur: String);
    procedure RechercheOnclick(Sender: TObject);
    procedure EnterAffaire(Sender: TObject);
    procedure EnterNumPhase(Sender: TObject);
    procedure EnterRessource(Sender: TObject);
    Procedure AffaireOnExit(Sender: TObject);

  end ;

Implementation

uses  AffaireUtil,
      AGLInitGC,
      FactUtil,
      SelectPhase;

procedure TOF_BTRECHPLANNING.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHPLANNING.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHPLANNING.OnUpdate ;
begin
  Inherited ;

  if ModePlanning <> 'PMA' then
    LaTOB.PutValue('RESSOURCE', Ressource.Text)
  else
    LaTOB.PutValue('MATERIEL', Ressource.Text);

  if (Affaire1.text = '') AND (Affaire2.text = '') AND (Affaire3.text = '') then Affaire.text := '';

  LaTOB.PutValue('AFFAIRE',   Affaire.Text);
  LaTOB.PutValue('NUMPHASE',  NumPhase.Text);

end ;

procedure TOF_BTRECHPLANNING.OnLoad ;
begin
  Inherited ;

  if LaTOB = nil then Exit;

  ModePlanning    := LaTOB.GetValue('MODEPLANNING');

  //
  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, nil, Tacreat, Affaire.text, True);

  if ModePlanning <> 'PMA' then
  Begin
    Ressource.Text:= LaTOB.GetValue('RESSOURCE');
    TRessource.Caption := 'Ressource :';
    if ModePlanning = 'PLA' then
    Begin
      Affaire0.text := 'W';
      TAffaire.Caption := 'Appel :';
    end
    else
    begin
      Affaire0.text := 'A';
      TAffaire.Caption := 'Chantier :';
    end;
  end
  else
  begin
    Ressource.Text:= LaTOB.GetValue('MATERIEL');
    TRessource.Caption  := 'Matériel :';
    Affaire0.text := '';
  end;

  Affaire.Text    := LaTOB.GetValue('AFFAIRE');
  NumPhase.Text   := LaTOB.GetValue('NUMPHASE');

end ;

procedure TOF_BTRECHPLANNING.OnArgument (S : String ) ;
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  NumPhase.Visible  := False;
  TNumPhase.Visible := False;

end ;

procedure TOF_BTRECHPLANNING.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHPLANNING.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHPLANNING.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BTRECHPLANNING.Controlechamp(Champ, Valeur : String);
begin
end;

procedure TOF_BTRECHPLANNING.GetObjects;
begin

  //
  Affaire       := THEdit(Getcontrol('AFFAIRE'));
  Affaire0      := THEdit(Getcontrol('AFFAIRE0'));
  Affaire1      := THEdit(Getcontrol('AFFAIRE1'));
  Affaire2      := THEdit(Getcontrol('AFFAIRE2'));
  Affaire3      := THEdit(Getcontrol('AFFAIRE3'));
  Avenant       := THEdit(Getcontrol('AVENANT'));
  //
  Ressource     := THEdit(Getcontrol('RESSOURCE'));
  NumPhase      := THEdit(Getcontrol('NUMPHASE'));
  BTRecherche   := TToolBarButton97(Getcontrol('BTRecherche'));
  //
  TRessource    := THLabel(GetControl('TRESSOURCE'));
  TAffaire      := THLabel(GetControl('TAFFAIRE'));
  TNumPhase     := THLabel(GetControl('TNUMPHASE'));
  //

end;

procedure TOF_BTRECHPLANNING.SetScreenEvents;
begin

  Affaire1.OnEnter          := EnterAffaire;
  Affaire2.OnEnter          := EnterAffaire;
  Affaire3.OnEnter          := EnterAffaire;

  Affaire1.OnExit           := AffaireOnExit;
  Affaire2.OnExit           := AffaireOnExit;
  Affaire3.OnExit           := AffaireOnExit;

  Ressource.OnEnter         := EnterRessource;
  NumPhase.OnEnter          := EnterNumPhase;


  BTRecherche.OnClick       := RechercheOnClick;

end;

Procedure TOF_BTRECHPLANNING.EnterAffaire(Sender : TObject);
Begin

  TRech := 'A';

end;

Procedure TOF_BTRECHPLANNING.EnterRessource(Sender : TObject);
Begin

  TRech := 'R';

end;

Procedure TOF_BTRECHPLANNING.EnterNumPhase(Sender : TObject);
Begin

  TRech := 'P';

end;


Procedure TOF_BTRECHPLANNING.RechercheOnclick(Sender : TObject);
Var StChamps : string;
    StWhere  : string;
    StSQL    : String;
begin

  if TRech = 'R' then
  begin
    StChamps  := ressource.Text;
    if ModePlanning <> 'PMA' then
    begin
      //StWhere :=  LaTOB.GetValue('TYPERESSOURCE');
      DispatchRecherche(Ressource, 3, stwhere, 'ARS_RESSOURCE=' + Trim(Ressource.Text)+';'+LaTOB.GetValue('TYPERESSOURCE'),'');
    end
    else
    begin
      StWhere :=  LaTOB.GetValue('TYPERESSOURCE');
      Ressource.text := AGLLanceFiche('BTP','BTMATERIEL_MUL','BMA_CODEMATERIEL='+StChamps+';'+LaTOB.GetValue('TYPERESSOURCE'), '', 'RECH=X');
    end;
    if ressource.text = '' then exit;
  end
  else if TRech = 'A' then
  begin
    StChamps  := Affaire.Text;
    if GetAffaireEnteteSt(Affaire0, Affaire1, Affaire2, Affaire3, AVENANT, nil, StChamps, false, false, false, True, true,'') then Affaire.text := StChamps;
    if Affaire.text = '' then Exit;
    ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, nil, TaModif, Affaire.text, False);
    if ModePlanning <> 'PMA' then
    begin
      if Affaire.Text <> '' then
      begin
        //Faudrait peut-être vérifier s'il y a des phase sur cette affaire !!!! 
        //Contrôle si des phase existent sur le chantier
        StSQL := 'SELECT GL_LIBELLE FROM LIGNE WHERE GL_NATUREPIECEG="PBT" AND GL_TYPELIGNE="DP1" AND GL_AFFAIRE="' + Affaire.Text + '"';
        If ExisteSQL(StSQL) then
        begin
          TNumPhase.Visible:= True;
          NumPhase.Visible := True;
          NumPhase.Plus := 'AND GL_AFFAIRE="' + Affaire.Text + '"';
        end;
      end;
    end;
  end
  else if TRech = 'P' then
  begin
    if Affaire.Text = '' then exit;
    if SelectionPhase (Affaire.text, StChamps) then
  	 NumPhase.Text := StChamps;
  end;

end;

Procedure TOF_BTRECHPLANNING.AffaireOnExit(Sender : Tobject);
Var CodeClient  : string;
    IP          : Integer;
Begin

  Affaire.text := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, CodeClient, TaModif, False, True, false, IP);

end;

Initialization
  registerclasses ( [ TOF_BTRECHPLANNING ] ) ;
end.

