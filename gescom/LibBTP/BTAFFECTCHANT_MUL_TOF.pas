{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/03/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BAFFECTCHANT_MUL ()
Mots clefs ... : TOF;BAFFECTCHANT_MUL
*****************************************************************}
Unit BTAFFECTCHANT_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     uTob,
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,     
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
     AGLInit,
     HTB97,
     uEntCommun,
     uTOFComm,
     UTOF;

Type
  TOF_BAFFECTCHANT_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
    Action        : TActionFiche;
    AppelPlanning : Boolean;
    //
    BOuvrir   : TToolbarButton97;
    BTSelect  : TToolbarButton97;

    BtEfface  : TToolbarButton97;
    //
    Affaire   : THEdit;
    Affaire0  : THEdit;
    Affaire1  : THEdit;
    Affaire2  : THEdit;
    Affaire3  : THEdit;
    Avenant   : THEdit;
    //
    LibAffaire: THLabel;
    //
    Grille    : THDBGrid;
    //
    procedure AffaireOnExit(Sender: TObject);
    procedure ChargeEnregEventInTob(IDAffect: integer);
    procedure ChargementTobPlanning;
    procedure Controlechamp(Critere, Champ, Valeur: String);
    procedure EffaceOnClick(Sender: TObject);
    procedure GetObjects;
    procedure GrilleOnDblclick(Sender: TObject);
    procedure OuvrirOnclick(Sender: TObject);
    procedure SelectOnClick(Sender: TObject);
    procedure SetScreenEvents;

  end ;

Implementation
uses  affaireutil,
      FactUtil,
      UtilsParc;

procedure TOF_BAFFECTCHANT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BAFFECTCHANT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BAFFECTCHANT_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BAFFECTCHANT_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BAFFECTCHANT_MUL.OnArgument (S : String ) ;
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin

  Inherited ;

  AppelPlanning := false;
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
     ControleChamp(Critere, Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.Text, False);

end ;

procedure TOF_BAFFECTCHANT_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BAFFECTCHANT_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BAFFECTCHANT_MUL.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BAFFECTCHANT_MUL.Controlechamp(Critere, Champ, Valeur : String);
begin

  if Champ = 'PLANNING' then AppelPlanning   := True
  else if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end


end;

Procedure TOF_BAFFECTCHANT_MUL.GetObjects;
begin

  if Assigned(GetControl('Fliste'))   then Grille    := THDBGrid(ecran.FindComponent('Fliste'));

  if Assigned(GetControl('BOuvrir'))  then BOuvrir   := TToolBarButton97(Getcontrol('BOuvrir'));

  if Assigned(GetControl('BSelect'))  then BTSelect  := TToolBarButton97(Getcontrol('BSELECT'));
  if Assigned(GetControl('BEfface'))  then BTEfface  := TToolBarButton97(Getcontrol('BEfface'));


  if Assigned(GetControl('BAT_AFFAIRE'))  then Affaire    := THEdit(GetControl('BAT_AFFAIRE'));
  if Assigned(GetControl('Affaire0')) then Affaire0       := THEdit(GetControl('AFFAIRE0'));
  if Assigned(GetControl('Affaire1')) then Affaire1       := THEdit(GetControl('AFFAIRE1'));
  if Assigned(GetControl('Affaire2')) then Affaire2       := THEdit(GetControl('AFFAIRE2'));
  if Assigned(GetControl('Affaire3')) then Affaire3       := THEdit(GetControl('AFFAIRE3'));
  if Assigned(GetControl('Avenant'))  then Avenant        := THEdit(GetControl('AVENANT'));

  if Assigned(GetControl('LibAffaire')) then LibAffaire   := THLabel(GetControl('LIBAFFAIRE'));

end;

procedure TOF_BAFFECTCHANT_MUL.SetScreenEvents;
begin

  If Assigned(Grille)                 then Grille.OnDblClick := GrilleOnDblclick;

  if Assigned(BOuvrir)                then BOuvrir.OnClick   := OuvrirOnClick;

  if Assigned(BTSelect)               then BTSelect.OnClick  := SelectOnClick;

  if Assigned(Affaire)                then Affaire.OnExit    := AffaireOnExit;

  if Assigned(BTEfface)               then BTEfface.OnClick  := EffaceOnClick;

end;

procedure TOF_BAFFECTCHANT_MUL.EffaceOnClick(Sender: TObject);
begin

  Affaire.Text  := '';
  Affaire0.Text := '';
  Affaire1.Text := '';
  Affaire2.Text := '';
  Affaire3.Text := '';
  Avenant.Text  := '';

  LibAffaire.caption := '...';

end;

procedure TOF_BAFFECTCHANT_MUL.AffaireOnExit(Sender : TObject);
Var IP      : Integer;
begin

  Affaire.text := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, '', Action, False, True, false, IP);

  LibAffaire.caption := ChargeInfoAffaire(Affaire.Text, False);

end;

procedure TOF_BAFFECTCHANT_MUL.SelectOnClick(Sender: TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire.Text;

  if GetAffaireEnteteSt(Affaire0, Affaire1, Affaire2, Affaire3, AVENANT, nil, StChamps, false, false, false, True, true,'') then Affaire.text := StChamps;

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.Text, False);

  LibAffaire.Caption := ChargeInfoAffaire(Affaire.Text, False);

end;

Procedure TOF_BAFFECTCHANT_MUL.OuvrirOnclick(Sender : TObject);
begin

  if AppelPlanning then
  begin
    ChargementTobPlanning;
    TFMUL(Ecran).Close;
  end;

end;

Procedure TOF_BAFFECTCHANT_MUL.GrilleOnDblclick(Sender : TObject);
Var Affaire : String;
    IDEvent : Integer;
begin

  Affaire :=  Grille.datasource.dataset.FindField('BAT_AFFAIRE').AsString;
  IdEvent :=  Grille.datasource.dataset.FindField('BAT_IDAFFECT').AsInteger;

  TFMul(Ecran).Retour := Affaire + ';' + IntToStr(IdEvent);
  TFMUL(Ecran).Close;

end;

procedure TOF_BAFFECTCHANT_MUL.ChargementTobPlanning;
Var QQ        : TQuery;
    IDAffect  : Integer;
    Ind       : Integer;
begin

  if Latob = nil then Exit;

  //chargement de la tob avec les lignes sélkectionnées
  if(Grille.NbSelected=0) and (not Grille.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné','');
    exit;
  end;

  SourisSablier;

  try
	  if Grille.AllSelected then
    begin
      QQ := TFMul(Ecran).Q;
      QQ.First;
      while Not QQ.EOF do
      Begin
        //on charge le Numéro d'évènement
        IDAffect := QQ.FindField('BAT_IDAFFECT').ASInteger;
        ChargeEnregEventInTob(IDAffect);
        QQ.Next;
      end;
      Grille.AllSelected:=False;
    end
    else
    begin
      for Ind :=0 to Grille.NbSelected-1 do
      begin
	      Grille.GotoLeBookmark(Ind);
        //On charge le N° d'évènement
        IDAffect := Grille.datasource.dataset.FindField('BAT_IDAFFECT').AsInteger;
        ChargeEnregEventInTob(IDAffect);
	    end;
      Grille.ClearSelected;
    end;
	finally
  	SourisNormale;
  end;

end;

Procedure TOF_BAFFECTCHANT_MUL.ChargeEnregEventInTob(IdAffect : integer);
var StSQL       : string;
    QQ          : TQuery;
    TobAffectPC : Tob;
begin

  StSQL := 'SELECT * FROM BAFFECTCHANT WHERE BAT_IDAFFECT=' + IntToStr(IdAffect);
  QQ    := OpenSQL(StSQL, False, -1,'',True);
  if not QQ.eof then
  begin
    TobAffectPC := TOB.Create('BAFFECTCHANT',LaTob,-1);
    TobAffectPC.SelectDB('',QQ);
  end;

  Ferme(QQ);

end;

Initialization
  registerclasses ( [ TOF_BAFFECTCHANT_MUL ] ) ;
end.

