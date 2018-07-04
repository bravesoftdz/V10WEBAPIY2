{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTYPEACTIONMAT_MUL ()
Mots clefs ... : TOF;BTTYPEACTIONMAT_MUL
*****************************************************************}
Unit BTTYPEACTIONMAT_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
     Fiche,
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
     uTob,
{$ELSE}
     DBCtrls, Db,
     {$IFNDEF DBXPRESS}
     dbTables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF DBXPRESS}
     fe_main,
     HDB,
     Mul,
{$ENDIF EAGLCLIENT}
     forms, 
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     //Ajout
     HTB97,
     uEntCommun,
     uTOFComm,
     UTOF;

Type
  TOF_BTTYPEACTIONMAT_MUL = Class (tTOFComm)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    //Définition des variables utilisées dans le Uses
    Action    : TActionFiche;
    //
    ModeRecherche : Boolean;
    //
    BInsert       : TToolbarButton97;
    BTSelect      : TToolbarButton97;
    BTSelect1      : TToolbarButton97;
    BTSelect2      : TToolbarButton97;
    BtEfface      : TToolbarButton97;
    BtEfface1      : TToolbarButton97;
    BtEfface2      : TToolbarButton97;
    //
    Grille        : THDBGrid;
    PComplement   : TTabSheet;
    //
    AssosDos      : TCheckBox;
    Modifiable    : TCheckBox;
    Valorise      : TCheckBox;
    GestionConso  : TCheckBox;
    //
    XX_WHERE      : THEdit;
    //
    Affaire       : THEdit;
    Affaire0      : THEdit;
    Affaire1      : THEdit;
    Affaire2      : THEdit;
    Affaire3      : THEdit;
    Avenant       : THEdit;
    //
    LibAffaire    : THLabel;
    CodeClient    : THEdit;
    LibClient     : THLabel;
    Ressource     : THEdit;
    LibRessource  : THLabel;
    //
    TypeAction    : String;
    //
    procedure Controlechamp(Champ, Valeur: String);
    procedure EffaceOnClick(Sender: TObject);
    procedure GetObjects;
    procedure GrilleOnDblclick(Sender: TObject);
    Procedure InsertOnclick(Sender : tobject);
    procedure SetScreenEvents;
    procedure SelectOnClick(Sender: TObject);
    procedure AffaireOnExit(Sender: TObject);
    procedure ClientOnExit(Sender: TObject);
    procedure Efface1OnClick(Sender: TObject);
    procedure Efface2OnClick(Sender: TObject);
    procedure RessourceOnExit(Sender: TObject);
    procedure Select1OnClick(Sender: TObject);
    procedure Select2OnClick(Sender: TObject);
  end ;

Implementation

uses  AffaireUtil,
      UtilsParc,
      FactUtil,
      AGLInitGC,
      BTPUtil;


procedure TOF_BTTYPEACTIONMAT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT_MUL.OnArgument (S : String ) ;
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
        Champ := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.Text, False);

  if TypeAction = 'PMA' then
  begin
    Ecran.caption := 'Recherche type d''action Parc/Matériel';
    PComplement.TabVisible:= True;
    AssosDos.Visible      := False;
    Modifiable.Visible    := False;
    Valorise.Visible      := True;
    GestionConso.Visible  := True;
  end
  else if TypeAction = 'PCA' then
  begin
    Ecran.caption := 'Recherche type d''action Chantier';
    PComplement.TabVisible:= False;
    AssosDos.Visible      := False;
    Modifiable.Visible    := False;
    Valorise.Visible      := False;
    GestionConso.Visible  := False;
  end
  else
  begin
    Ecran.caption := 'Recherche type d''action Intervention';
    PComplement.TabVisible:= False;
    AssosDos.Visible      := True;
    Modifiable.Visible    := True;
    Valorise.Visible      := False;
    GestionConso.Visible  := False;
  end;

  UpdateCaption(Ecran);

end ;

procedure TOF_BTTYPEACTIONMAT_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT_MUL.GetObjects;
begin

  if Assigned(GetControl('Fliste'))   then Grille   := THDBGrid(ecran.FindComponent('Fliste'));

  if Assigned(GetControl('BInsert'))  then BInsert  := TToolBarButton97(Getcontrol('BInsert'));

  if Assigned(GetControl('BSelect'))          then BTSelect  := TToolBarButton97(Getcontrol('BSELECT'));
  if Assigned(GetControl('BSelect1'))         then BTSelect1 := TToolBarButton97(Getcontrol('BSELECT1'));
  if Assigned(GetControl('BSelect2'))         then BTSelect2 := TToolBarButton97(Getcontrol('BSELECT2'));

  Affaire     := THEdit(GetControl('BTA_AFFAIRE'));
  Affaire0    := THEdit(GetControl('BAM_AFFAIRE0'));
  Affaire1    := THEdit(GetControl('BAM_AFFAIRE1'));
  Affaire2    := THEdit(GetControl('BAM_AFFAIRE2'));
  Affaire3    := THEdit(GetControl('BAM_AFFAIRE3'));
  Avenant     := THEdit(GetControl('BAM_AVENANT'));

  if Assigned(GetControl('BTA_TIERS'))        then CodeClient:= THEdit(GetControl('BTA_TIERS'));
  if Assigned(GetControl('BTA_RESSOURCE'))    then Ressource := THEdit(GetControl('BTA_RESSOURCE'));

  if Assigned(GetControl('TAFFAIRE'))         then LibAffaire   := THLabel(GetControl('TAFFAIRE'));
  if Assigned(GetControl('TCLIENT'))          then LibClient    := THLabel(GetControl('TCLIENT'));
  if Assigned(GetControl('TRESOSURCE'))       then LibRessource := THLabel(GetControl('TRESSOURCE'));

  if Assigned(GetControl('BEfface'))          then BTEfface  := TToolBarButton97(Getcontrol('BEFFACE'));
  if Assigned(GetControl('BEfface1'))         then BTEfface1 := TToolBarButton97(Getcontrol('BEFFACE1'));
  if Assigned(GetControl('BEfface2'))         then BTEfface2 := TToolBarButton97(Getcontrol('BEFFACE2'));

  if Assigned(GetControl('PComplement'))      then PComplement  := TTabSheet(GetControl('PComplement'));

  if Assigned(GetControl('BTA_ASSOSDOS'))     then Assosdos     := TCheckBox(GetControl('BTA_ASSOSDOS'));
  if Assigned(GetControl('BTA_MODIFIABLE'))   then Modifiable   := TCheckBox(GetControl('BTA_MODIFIABLE'));
  if Assigned(GetControl('BTA_VALORISE'))     then Valorise     := TCheckBox(GetControl('BTA_VALORISE'));
  if Assigned(GetControl('BTA_GESTIONCONSO')) Then GestionConso := TCheckBox(GetControl('BTA_GESTIONCONSO'));

  if Assigned(GetControl('XX_WHERE')) then XX_WHERE := THEdit(GetControl('XX_WHERE'));

end;
procedure TOF_BTTYPEACTIONMAT_MUL.SetScreenEvents;
begin

  If Assigned(Grille)                 then Grille.OnDblClick := GrilleOnDblclick;

  if Assigned(BInsert)                then BInsert.OnClick   := InsertOnClick;
  if Assigned(BTSelect)               then BTSelect.OnClick  := SelectOnClick;
  if Assigned(BTEfface)               then BTEfface.OnClick  := EffaceOnClick;

  if Assigned(BTSelect1)              then BTSelect1.OnClick := Select1OnClick;
  if Assigned(BTSelect2)              then BTSelect2.OnClick := Select2OnClick;

  if Assigned(Ressource)              then Ressource.OnExit  := RessourceOnExit;
  if Assigned(Affaire)                then Affaire.OnExit    := AffaireOnExit;
  if Assigned(CodeClient)             Then CodeClient.OnExit := ClientOnExit;

  if Assigned(BTEfface1)              then BTEfface1.OnClick := Efface1OnClick;
  if Assigned(BTEfface2)              then BTEfface2.OnClick := Efface2OnClick;

end;

Procedure TOF_BTTYPEACTIONMAT_MUL.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if Valeur='CREATION'          then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end
  else if Champ = 'TYPEACTION'    then
  begin
    TypeAction := Valeur;
    XX_WHERE.Text := ' AND BTA_TYPEACTION="' + TypeAction + '"';
  end
  else if Champ = 'RECH' then
  begin
    if valeur='X'                 then ModeRecherche:=true;
  end;

end;

Procedure TOF_BTTYPEACTIONMAT_MUL.GrilleOnDblclick(Sender : TObject);
Var Argument  : String;
    BtEtat    : String;
begin

  BtEtat := Grille.Datasource.Dataset.FindField('BTA_BTETAT').AsString;

  if Moderecherche then
  begin
  	TFMul(Ecran).Retour := BtEtat;
    TFMul(Ecran).Close;
  end
  else
  begin
    if BtEtat = '' then
      Argument := 'ACTION=CREATION'
    else
      Argument := 'TYPEACTION=' + TypeAction + ';BTETAT=' + BtEtat + ';ACTION=MODIFICATION';

    AGLLanceFiche('BTP','BTTYPEACTIONMAT','','',Argument);

    refreshDB();
  end;

end;

procedure TOF_BTTYPEACTIONMAT_MUL.InsertOnclick(Sender: TObject);
Var Argument : String;
begin

  Argument := 'TYPEACTION=' + TypeAction + ';ACTION=CREATION';

  AGLLanceFiche('BTP','BTTYPEACTIONMAT','','',Argument);

  refreshDB();

end;

procedure TOF_BTTYPEACTIONMAT_MUL.SelectOnClick(Sender: TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire.Text;

  if GetAffaireEnteteSt(Affaire0, Affaire1, Affaire2, Affaire3, AVENANT, nil, StChamps, false, false, false, True, true,'') then Affaire.text := StChamps;

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.Text, False);

end;

procedure TOF_BTTYPEACTIONMAT_MUL.EffaceOnClick(Sender: TObject);
begin

  Affaire.Text  := '';
  Affaire0.Text := '';
  Affaire1.Text := '';
  Affaire2.Text := '';
  Affaire3.Text := '';
  Avenant.Text  := '';

end;

procedure TOF_BTTYPEACTIONMAT_MUL.AffaireOnExit(Sender : TObject);
Var IP      : Integer;
begin

  Affaire.text := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, '', Action, False, True, false, IP);


  LibAffaire.caption := ChargeInfoAffaire(Affaire.Text, False);

end;

procedure TOF_BTTYPEACTIONMAT_MUL.Select1OnClick(Sender: TObject);
Var StChamps  : string;
    StWhere   : string;
Begin

  StChamps  := CodeClient.Text;

  StWhere   := 'T_NATUREAUXI="CLI"';

  DispatchRecherche(CodeClient, 2, StWhere,'T_TIERS=' + Trim(CodeClient.Text), '');

  LibClient.caption := ChargeInfoTiers(CodeClient.Text, False);

end;

procedure TOF_BTTYPEACTIONMAT_MUL.ClientOnExit(Sender : TObject);
begin

  LibClient.caption := ChargeInfoTiers(CodeClient.Text, False);

end;

procedure TOF_BTTYPEACTIONMAT_MUL.RessourceOnExit(Sender : TObject);
Var StWhere : string;
begin

  Stwhere   := ' AND ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")';

  LibRessource.Caption := ChargeInfoRessource(Ressource.Text, StWhere, False);

end;

procedure TOF_BTTYPEACTIONMAT_MUL.Select2OnClick(Sender: TObject);
Var StChamps  : String;
    StWhere   : String;
begin

  StChamps  := ressource.Text;

  Stwhere   := 'ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")';

  DispatchRecherche(Ressource, 3, StWhere,'ARS_RESSOURCE=' + Trim(Ressource.Text), '');

  LibRessource.Caption := ChargeInfoRessource(Ressource.Text, ' AND ' + StWhere, False);

end;

procedure TOF_BTTYPEACTIONMAT_MUL.Efface1OnClick(Sender: TObject);
begin

  CodeClient.text := '';
  LibClient.Caption := '';

end;

procedure TOF_BTTYPEACTIONMAT_MUL.Efface2OnClick(Sender: TObject);
begin

  Ressource.text := '';
  LibRessource.caption := '';

end;

Initialization
  registerclasses ( [ TOF_BTTYPEACTIONMAT_MUL ] ) ;
end.

