{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTYPEACTIONMAT_MUL ()
Mots clefs ... : TOF;BTTYPEACTIONMAT_MUL
*****************************************************************}
Unit BTAFFECTATION_MUL_TOF ;

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
  TOF_BTAFFECTATION_MUL = Class (tTOFComm)
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
    Action        : TActionFiche;
    //
    BInsert       : TToolbarButton97;
    Grille        : THDBGrid;
    //
    BTSelect1 : TToolbarButton97;
    BTSelect2 : TToolbarButton97;
    BTSelect3 : TToolbarButton97;

    BtEfface1 : TToolbarButton97;
    BtEfface2 : TToolbarButton97;
    BtEfface3 : TToolbarButton97;
    //
    StFamilleMat  : THLabel;
    StMateriel    : THLabel;
    StRessource   : THLabel;
    //
    FamilleMat    : THEdit;
    CodeMateriel  : THEdit;
    Ressource     : THEdit;
    //
    procedure Controlechamp(Champ, Valeur: String);

    procedure Efface1OnClick(Sender: TObject);
    procedure Efface2OnClick(Sender: TObject);
    procedure Efface3OnClick(Sender: TObject);

    procedure FamilleMatOnExit(Sender: TObject);

    procedure GetObjects;
    procedure GrilleOnDblclick(Sender: TObject);

    Procedure InsertOnclick(Sender : tobject);

    procedure MaterielOnExit(Sender: TObject);
    procedure RessourceOnExit(Sender: TObject);

    procedure Select1OnClick(Sender: TObject);
    procedure Select2OnClick(Sender: TObject);
    procedure Select3OnClick(Sender: TObject);
    procedure SetScreenEvents;
  end ;

Implementation

uses  AffaireUtil,
      UtilsParc,
      AGLInitGC,
      UtofRessource_Mul,
      FactUtil;


procedure TOF_BTAFFECTATION_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFECTATION_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFECTATION_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFECTATION_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFECTATION_MUL.OnArgument (S : String ) ;
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

procedure TOF_BTAFFECTATION_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFECTATION_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFECTATION_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFECTATION_MUL.GetObjects;
begin

  if Assigned(GetControl('Fliste'))   then Grille   := THDBGrid(ecran.FindComponent('Fliste'));

  if Assigned(GetControl('BInsert'))  then BInsert  := TToolBarButton97(Getcontrol('BInsert'));

  if Assigned(GetControl('BFm_CODEFAMILLE'))   then FamilleMat   := THEdit(GetControl('BFM_CODEFAMILLE'));
  if Assigned(GetControl('BFF_CODEMATERIEL'))  then CodeMateriel := THEdit(GetControl('BFF_CODEMATERIEL'));
  if Assigned(GetControl('BFF_RESSOURCE'))     then Ressource    := THEdit(GetControl('BFF_RESSOURCE'));

  if Assigned(GetControl('STFAMILLEMAT'))      then StFamilleMat := THLabel(GetControl('STFAMILLEMAT'));
  if Assigned(GetControl('STMATERIEL'))        then StMateriel   := THLabel(GetControl('STMATERIEL'));
  if Assigned(GetControl('STRESSOURCE'))       then StRessource  := THLabel(GetControl('STRESSOURCE'));

  if Assigned(GetControl('BSelect1')) then BTSelect1 := TToolBarButton97(Getcontrol('BSELECT1'));
  if Assigned(GetControl('BSelect2')) then BTSelect2 := TToolBarButton97(Getcontrol('BSELECT2'));
  if Assigned(GetControl('BSelect3')) then BTSelect3 := TToolBarButton97(Getcontrol('BSELECT3'));

  if Assigned(GetControl('BEfface1')) then BTEfface1 := TToolBarButton97(Getcontrol('BEfface1'));
  if Assigned(GetControl('BEfface2')) then BTEfface2 := TToolBarButton97(Getcontrol('BEfface2'));
  if Assigned(GetControl('BEfface3')) then BTEfface3 := TToolBarButton97(Getcontrol('BEfface3'));


end;
procedure TOF_BTAFFECTATION_MUL.SetScreenEvents;
begin

  If Assigned(Grille)       then Grille.OnDblClick := GrilleOnDblclick;

  if Assigned(BTSelect1)    then BTSelect1.OnClick := Select1OnClick;
  if Assigned(BTSelect2)    then BTSelect2.OnClick := Select2OnClick;
  if Assigned(BTSelect3)    then BTSelect3.OnClick := Select3OnClick;

  if Assigned(BTEfface1)    then BTEfface1.OnClick := Efface1OnClick;
  if Assigned(BTEfface2)    then BTEfface2.OnClick := Efface2OnClick;
  if Assigned(BTEfface3)    then BTEfface3.OnClick := Efface3OnClick;

  if Assigned(BInsert)      then BInsert.OnClick   := InsertOnClick;

  If Assigned(FamilleMat)   then FamilleMat.OnExit    := FamilleMatOnExit;
  If Assigned(CodeMateriel) then CodeMateriel.OnExit  := MaterielOnExit;
  if Assigned(Ressource)    then Ressource.OnExit     := RessourceOnExit;


end;

Procedure TOF_BTAFFECTATION_MUL.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if Valeur='CREATION'          then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

end;

procedure TOF_BTAFFECTATION_MUL.FamilleMatOnExit(Sender: TObject);
begin

  StFamilleMat.caption := ChargeInfoFamMat(FamilleMat.Text, False);

end;

procedure TOF_BTAFFECTATION_MUL.MaterielOnExit(Sender : TObject);
begin

  StMateriel.caption := ChargeinfoMateriel(CodeMateriel.text, True);

end;

procedure TOF_BTAFFECTATION_MUL.RessourceOnExit(Sender : TObject);
Var StWhere : String;
begin

  StWhere := ' AND ARS_TYPERESSOURCE ="SAL"';

  StRessource.caption := ChargeinfoRessource(Ressource.text, StWhere, True);

end;

Procedure TOF_BTAFFECTATION_MUL.GrilleOnDblclick(Sender : TObject);
Var Argument    : String;
    Identification  : String;
begin

  Identification := Grille.Datasource.Dataset.FindField('BFF_IDAFFECTATION').AsString;

  if Identification = '' then
    Argument := 'ACTION=CREATION'
  else
    Argument := 'IDAFFECTATION=' + Identification + ';ACTION=MODIFICATION';

  AGLLanceFiche('BTP','BTAFFECTPARC','','',Argument);

  refreshDB();

end;

procedure TOF_BTAFFECTATION_MUL.InsertOnclick(Sender: TObject);
Var Argument : String;
begin

  Argument := 'ACTION=CREATION';

  AGLLanceFiche('BTP','BTAFFECTPARC','','',Argument);

  refreshDB();

end;

procedure TOF_BTAFFECTATION_MUL.Select1OnClick(Sender: TObject);
begin

  FamilleMat.text := AGLLanceFiche('BTP','BTFAMILLEMAT_MUL','','','RECH=X');

  StFamilleMat.caption := ChargeInfoFamMat(FamilleMat.Text, False);

end;

procedure TOF_BTAFFECTATION_MUL.Select2OnClick(Sender: TObject);
begin

  CodeMateriel.text := AGLLanceFiche('BTP','BTMATERIEL_MUL','','','RECH=X');

  StMateriel.caption := ChargeInfoMateriel(CodeMateriel.Text, False);

end;

Procedure TOF_BTAFFECTATION_MUL.Select3OnClick(Sender : TObject);
Var StChamps  : string;
    StWhere   : string;
    Action    : string;
begin

  StChamps  := ressource.Text;

  StWhere := 'TYPERESSOURCE=SAL';
 	Action := ';ACTION=RECH';

  Ressource.Text := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+Trim(StChamps),stwhere + Action);
  StRessource.caption := ChargeInfoRessource(StChamps, ' AND ' + StWhere, True);


end;

procedure TOF_BTAFFECTATION_MUL.Efface1OnClick(Sender: TObject);
begin

  FamilleMat.Text := '';
  StFamilleMat.caption := '';

end;

procedure TOF_BTAFFECTATION_MUL.Efface2OnClick(Sender: TObject);
begin

  CodeMateriel.Text := '';
  StMateriel.Caption := '';

end;

procedure TOF_BTAFFECTATION_MUL.Efface3OnClick(Sender: TObject);
begin

  Ressource.text := '';
  StRessource.caption := '';

end;

Initialization
  registerclasses ( [ TOF_BTAFFECTATION_MUL ] ) ;
end.

