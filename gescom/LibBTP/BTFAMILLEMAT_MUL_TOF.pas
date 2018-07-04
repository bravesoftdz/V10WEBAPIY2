{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFAMILLEMAT_MUL ()
Mots clefs ... : TOF;BTFAMILLEMAT_MUL
*****************************************************************}
Unit BTFAMILLEMAT_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
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
  TOF_BTFAMILLEMAT_MUL = Class (tTOFComm)
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
    Moderecherche : Boolean;
    //
    BInsert   : TToolbarButton97;
    Grille    : THDBGrid;
    //
    procedure Controlechamp(Champ, Valeur: String);
    procedure GetObjects;
    procedure GrilleOnDblclick(Sender: TObject);
    Procedure InsertOnclick(Sender : tobject);
    procedure SetScreenEvents;
  end ;

Implementation


procedure TOF_BTFAMILLEMAT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMAT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMAT_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMAT_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMAT_MUL.OnArgument (S : String ) ;
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


end ;

procedure TOF_BTFAMILLEMAT_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMAT_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMAT_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTFAMILLEMAT_MUL.GetObjects;
begin

  if Assigned(GetControl('Fliste'))   then Grille   := THDBGrid(ecran.FindComponent('Fliste'));

  if Assigned(GetControl('BInsert'))  then BInsert := TToolBarButton97(Getcontrol('BInsert'));

end;
procedure TOF_BTFAMILLEMAT_MUL.SetScreenEvents;
begin

  Grille.OnDblClick := GrilleOnDblclick;

  BInsert.OnClick   := InsertOnClick;

end;

Procedure TOF_BTFAMILLEMAT_MUL.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if Valeur='CREATION'          then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end
  else if Champ = 'RECH' then
  begin
    if valeur='X'                 then ModeRecherche:=true;
  end;

end;

Procedure TOF_BTFAMILLEMAT_MUL.GrilleOnDblclick(Sender : TObject);
Var Argument    : String;
    FamilleMat  : String;
begin

  FamilleMat := Grille.Datasource.Dataset.FindField('BFM_CODEFAMILLE').AsString;

  if Moderecherche then
  begin
  	TFMul(Ecran).Retour := FamilleMat;
    TFMul(Ecran).Close;
  end
  else
  begin
    if FamilleMat = '' then
      Argument := 'ACTION=CREATION'
    else
      Argument := 'CODEFAMILLE=' + FamilleMat + ';ACTION=MODIFICATION';

    AGLLanceFiche('BTP','BTFAMILLEMATERIEL','','',Argument);

    refreshDB();
  end;

end;

procedure TOF_BTFAMILLEMAT_MUL.InsertOnclick(Sender: TObject);
Var Argument : String;
begin

  Argument := 'ACTION=CREATION';

  AGLLanceFiche('BTP','BTFAMILLEMATERIEL','','',Argument);

  refreshDB();

end;

Initialization
  registerclasses ( [ TOF_BTFAMILLEMAT_MUL ] ) ;
end.

