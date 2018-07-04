{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFAMILLEMAT_MUL ()
Mots clefs ... : TOF;BTFAMILLEMAT_MUL
*****************************************************************}
Unit BTMATERIEL_MUL_TOF ;

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
  TOF_BTMATERIEL_MUL = Class (tTOFComm)
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
    ModeRecherche : Boolean;
    //
    CodeFamille : THEdit;
    Ressource   : THEdit;
    CodeArticle : THEdit;
    //
    StRessource : THLabel;
    StFamille   : THLabel;
    StArticle   : THLabel;
    //
    BInsert   : TToolbarButton97;
    Grille    : THDBGrid;
    //
    procedure ArticleOnElipsisClick(Sender: Tobject);
    procedure Controlechamp(Champ, Valeur: String);
    procedure GetObjects;
    procedure GrilleOnDblclick(Sender: TObject);
    Procedure InsertOnclick(Sender : tobject);
    procedure SetScreenEvents;
    procedure ArticleOnExit(Sender: TObject);
    procedure FamilleOnExit(Sender: TObject);
    procedure RessourceOnExit(Sender: TObject);
    procedure RessourceOnElipsisClick(Sender: Tobject);
  end ;

Implementation

  Uses UtilsParc,
       LookUp;


procedure TOF_BTMATERIEL_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL_MUL.OnArgument (S : String ) ;
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
        Champ  := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;   

end ;

procedure TOF_BTMATERIEL_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL_MUL.GetObjects;
begin

  if Assigned(GetControl('Fliste'))           then Grille   := THDBGrid(ecran.FindComponent('Fliste'));

  if Assigned(GetControl('BInsert'))          then BInsert := TToolBarButton97(Getcontrol('BInsert'));

  If Assigned(GetControl('BMA_CODEFAMILLE'))  then CodeFamille := THEdit(GetCOntrol('BMA_CODEFAMILLE'));
  If Assigned(GetControl('BMA_RESSOURCE'))    then Ressource   := THEdit(GetCOntrol('BMA_RESSOURCE'));
  If Assigned(GetControl('BMA_CODEARTICLE'))  then CodeArticle := THEdit(GetCOntrol('BMA_CODEARTICLE'));

  if Assigned(GetCONTROL('LIBFAMILLEMAT'))    then StFamille   := THLabel(GetCONTROL('LIBFAMILLEMAT'));
  if Assigned(GetCONTROL('LIBRESSOURCE'))     then StRessource := THLabel(GetCONTROL('LIBRESSOURCE'));
  if Assigned(GetCONTROL('LIBARTICLE'))       then StArticle   := THLabel(GetCONTROL('LIBARTICLE'));


end;
procedure TOF_BTMATERIEL_MUL.SetScreenEvents;
begin

  Grille.OnDblClick         := GrilleOnDblclick;

  BInsert.OnClick           := InsertOnClick;

  CodeFamille.OnExit        := FamilleOnExit;
  Ressource.OnExit          := RessourceOnExit;
  Ressource.OnElipsisClick  := RessourceOnElipsisClick;
  CodeArticle.OnExit        := ArticleOnExit;
  CodeArticle.OnElipsisClick:= ArticleOnElipsisClick;

end;

Procedure TOF_BTMATERIEL_MUL.Controlechamp(Champ, Valeur : String);
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

Procedure TOF_BTMATERIEL_MUL.GrilleOnDblclick(Sender : TObject);
Var Argument : String;
    CodeMat  : string;
begin

  CodeMat := Grille.Datasource.Dataset.FindField('BMA_CODEMATERIEL').AsString;

  if Moderecherche then
  begin
  	TFMul(Ecran).Retour := CodeMat;
    TFMul(Ecran).Close;
  end
  else
  begin
    if CodeMat = '' then
      Argument := 'ACTION=CREATION'
    else
      Argument := 'CODEMATERIEL=' + CodeMat + ';ACTION=MODIFICATION';

      AGLLanceFiche('BTP','BTMATERIEL','','',Argument);

    refreshDB();
  end;


end;

procedure TOF_BTMATERIEL_MUL.InsertOnclick(Sender: TObject);
Var Argument : String;
begin

  Argument := 'ACTION=CREATION';

  AGLLanceFiche('BTP','BTMATERIEL','','',Argument);

  refreshDB();

end;

procedure TOF_BTMATERIEL_MUL.FamilleOnExit(Sender: TObject);
begin

  if codeFamille.text <> '' then Stfamille.caption := ChargeInfoFamMat(CodeFamille.Text,False);

end;

procedure TOF_BTMATERIEL_MUL.RessourceOnExit(Sender: TObject);
Var Stwhere : string;
begin

  stWhere := 'AND ARS_TYPERESSOURCE IN ("OUT","MAT","LOC","AUT")';

  if Ressource.text <> '' then StRessource.caption := ChargeInfoRessource(Ressource.Text, StWhere, False);

end;


procedure TOF_BTMATERIEL_MUL.RessourceOnElipsisClick(Sender: Tobject);
Var Title   : string;
    StWhere : string;
begin

  title := 'Recherche Ressource Matériel';

  stWhere := 'AND ARS_TYPERESSOURCE IN ("OUT","MAT","LOC","AUT")';

  if not LookupList(Ressource,Title,'RESSOURCE','ARS_RESSOURCE','ARS_LIBELLE',StWhere,'ARS_LIBELLE',True,7) then Exit;

  StRessource.Caption := ChargeInfoRessource(Ressource.Text, Stwhere, False);

end;

procedure TOF_BTMATERIEL_MUL.ArticleOnElipsisClick(Sender: Tobject);
Var StWhere : String;
    StChamps: String;
    title   : string;
begin

  title := 'Recherche Prestation Associée';

  stWhere := '(GA_TYPEARTICLE="PRE")';

  stchamps := 'TYPERESSOURCE=MAT,LOC,OUT';

  if stWhere <> '' then StChamps := StChamps + ';' + stWhere;

  CodeArticle.text := AGLLanceFiche('BTP', 'BTPREST_RECH', 'GA_CODEARTICLE='+Trim(Copy(CodeArticle.text, 1, 18)), '', stchamps);


  if CodeArticle.text <> '' then
  begin
    StWhere := ' AND GA_TYPEARTICLE="PRE"';
    CodeArticle.text := Trim(Copy(CodeArticle.text, 1, 18));
    StArticle.caption := ChargeInfoArticle(CodeArticle.Text, StWhere, False);
  end;

end;


procedure TOF_BTMATERIEL_MUL.ArticleOnExit(Sender: TObject);
Var StWhere : string;
begin

  StWhere := ' AND GA_TYPEARTICLE="PRE"';

  if CodeArticle.text <> '' then StArticle.caption := ChargeInfoArticle(CodeArticle.Text, StWhere, False);

end;

Initialization
  registerclasses ( [ TOF_BTMATERIEL_MUL ] ) ;
end.

