{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 10/03/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCODEBARRE ()
Mots clefs ... : TOF;BTCODEBARRE
*****************************************************************}
Unit Utof_BTCODEBARRE;

Interface

Uses vierge,
     StdCtrls,
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
     uTOB,
     HTB97,
     UTOF ;
Type
  TOF_BTCODEBARRE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    //
  Private
    //Définition des variables utilisées dans le Uses
    Action        : TActionFiche;
    //
    Ok_CtrlSaisie : Boolean;
    //
    NatureCAB     : THEdit;
    IdentifCAB    : THEdit;
    CodeBarre     : THEdit;
    //
    NomID         : THLabel;
    //
    Binsert       : TToolBarButton97;
    BValider      : TToolBarButton97;
    BDelete       : TToolBarButton97;
    //
    CabPrincipal  : THCheckbox;
    QualifCAB     : THValComboBox;
    //
    Grille        : THGrid;
    //
    fColNamesGS   : string;
    Falignement   : string;
    Ftitre        : string;
    fLargeur      : string;
    //
    Ok_Action     : String;
    //
    QQ            : TQuery ;
    StSQL         : String ;
    //
    TOBCodeBarre  : TOB;
    TOBL          : TOB;

    procedure ChargeGrilleToScreen(Arow: Integer);
    procedure ChargeLibAffaire;
    procedure ChargeLibArticle;
    procedure ChargeLibFournisseur;
    procedure ChargementGrille;
    procedure ChargeScreenToGrille;
    procedure ChargeTOB;
    procedure CodeBarreOnExit(Sender: Tobject);
    Function  ControlCAB: Integer;
    procedure Controlechamp(Champ, Valeur: String);
    procedure CreateTOB;
    procedure DefinieGrid;
    procedure DeleteOnClick(Sender: Tobject);
    procedure DessineGrille(GS: THGRID);
    procedure DestroyTOB;
    procedure GetObjects;
    procedure GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure InsertOnClick(Sender: Tobject);
    procedure PrincipalOnClick(Sender: Tobject);
    procedure SetGrilleEvents(Etat: boolean);
    procedure SetScreenEvents;
    procedure ValideOnClick(Sender: Tobject);

    end ;

const
	// libellés des messages
  TexteMessage: array[1..7] of string  = (
          {1}  'Le code barre et le qualifiant code barre doivent être renseignés  '
          {2}, 'Le nombre de caractères est incorrect'
          {3}, 'Le code barre est incorrect'
          {4}, 'Le code barre contient des caractères non autorisés'
          {5}, 'Le code barre doit être numérique'
          {6}, 'Confirmez-vous la suppression de cette ligne ?'
          {7}, 'Vous ne pouvez avoir qu''un seul code barre principal !'
                                         );

Implementation
Uses UtilArticle;


procedure TOF_BTCODEBARRE.OnArgument ( S: String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin

  Inherited ;
  //
  Ok_CtrlSaisie := False;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  Critere := S;
  //
  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Gestion des évènement de Grille
  SetGrilleEvents(True);

  // définition de la grille
  DefinieGrid;

  //dessin de la grille
  DessineGrille(Grille);

  //Chargement de la Tob en fonction des élément préchargés.
  ChargeTOB;

  //chargement de la grille avec les éléments de la TOB
  ChargementGrille;

  //Chargement de la premièReal ligne
  if TOBCodeBARRE.detail.count <> 0 then ChargeGrilleToScreen(1);

  if NatureCAB.Text = 'FOU' then
    ChargeLibFournisseur
  else if NatureCAB.Text = 'MAR' then
    ChargeLibArticle
  else if NatureCAB.Text = 'AFF' then
    ChargeLibAffaire
  else NomID.caption := '';

  NatureCAB.Enabled := False;
  IdentifCAB.enabled:= False;

  if Action = taConsult then
  begin
    Binsert.Visible := false;
    BDelete.Visible := False;
    BValider.Visible:= False;
  end;

end ;

Procedure TOF_BTCODEBARRE.ChargeLibFournisseur;
begin

  StSQL := 'SELECT T_LIBELLE FROM TIERS ';
  StSQL := StSQL + ' WHERE T_NATUREAUXI="' + NatureCab.text + '" ';
  StSQL := StSQL + ' AND   T_TIERS="' + IdentifCab.Text + '" ';

  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then NomID.caption := QQ.Findfield('T_LIBELLE').AsString;

  Ferme(QQ) ;

end;

Procedure TOF_BTCODEBARRE.ChargeLibArticle;
begin

  StSQL := 'SELECT GA_LIBELLE FROM ARTICLE ';
  StSQL := StSQL + ' WHERE GA_TYPEARTICLE="' + NatureCab.text + '" ';
  StSQL := StSQL + ' AND   GA_ARTICLE="' + IdentifCab.Text + '" ';

  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then NomID.caption := QQ.Findfield('GA_LIBELLE').AsString;

  Ferme(QQ) ;

end;

Procedure TOF_BTCODEBARRE.ChargeLibAffaire;
begin

  StSQL := 'SELECT AFF_LIBELLE FROM AFFAIRE ';
  StSQL := StSQL + ' WHERE AFF_STATUTAFFAIRE="' + NatureCab.text + '" ';
  StSQL := StSQL + ' AND   AFF_AFFAIRE="' + IdentifCab.Text + '" ';

  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then NomID.caption := QQ.Findfield('AFF_LIBELLE').AsString;

  Ferme(QQ) ;

end;


procedure TOF_BTCODEBARRE.CreateTOB;
begin

  TOBCodeBarre := TOB.Create('LES CODEBARRE',nil,-1);

end;

procedure TOF_BTCODEBARRE.DestroyTOB;
begin

  FreeAndNil(TOBCodeBarre)

end;

Procedure TOF_BTCODEBARRE.ChargeTOB;
begin

  TobCodeBarre.ClearDetail;

  StSQL := 'SELECT * FROM BTCODEBARRE ';
  StSQL := StSQL + ' WHERE BCB_NATURECAB = "' + NatureCab.text  + '" ';
  StSQL := StSQL + ' AND   BCB_IDENTIFCAB= "' + IdentifCAB.text + '" ';

  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then TOBCodeBarre.LoadDetailDB('BTCODEBARRE','','',QQ,false,true) ;

  Ferme(QQ) ;

end; 

procedure TOF_BTCODEBARRE.DefinieGrid;
begin

  // Définition de la liste de saisie pour la grille Détail
  fColNamesGS := 'SEL;CC_QUALIFCAB;CC_CODEBARRE';
  Falignement := 'C.0  ---;C.0  ---;G.0  ---';
  Ftitre      := ' ;Type;Code Barre';
  fLargeur    := '2;4;15';

end;

procedure TOF_BTCODEBARRE.DessineGrille(GS : THGRID);
var st              : String;
    lestitres       : String;
    lesalignements  : String;
    FF              : String;
    alignement      : String;
    Nam             : String;
    leslargeurs     : String;
    lalargeur       : String;
    letitre         : String;
    lelement        : string;
    //
    Obli            : Boolean;
    OkLib           : Boolean;
    OkVisu          : Boolean;
    OkNulle         : Boolean;
    OkCumul         : Boolean;
    Sep             : Boolean;
    Okimg           : boolean;
    //
    dec             : Integer;
    NbCols          : integer;
    indice          : Integer;
begin
  //
  //Calcul du nombre de Colonnes du Tableau en fonction des noms
  st := fColNamesGS;

  NbCols := 0;

  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  GS.ColCount     := Nbcols;
  //
  st              := fColNamesGS ;
  lesalignements  := Falignement;
  lestitres       := Ftitre;
  leslargeurs     := Flargeur;

  //Mise en forme des colonnes
  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    //
    alignement  := ReadTokenSt(lesalignements);
    lalargeur   := readtokenst(leslargeurs);
    letitre     := readtokenst(lestitres);
    //
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    //
    GS.cells[Indice,0]     := leTitre;
    GS.ColNames [Indice]   := Nam;
    //
    //Alignement des cellules
    if copy(Alignement,1,1)='G'       then //Cadré à Gauche
      GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D'  then  //Cadré à Droite
      GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C'  then
      GS.ColAligns[indice] := taCenter; //Cadré au centre

    //Colonne visible ou non
    if OkVisu then
  		GS.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W')
    else
    	GS.ColWidths[indice] := -1;

    //Affichage d'une image ou du texte
    okImg := (copy(Alignement,8,1)='X');
    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	GS.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end;

    if (Dec<>0) or (Sep) then
    begin
    	if OkNulle then
        GS.ColFormats[indice] := FF+';; ;' //'#'
      else
      	GS.ColFormats[indice] := FF; //'#';
    end;
  end ;

end;

Procedure TOF_BTCODEBARRE.ChargementGrille;
var indice : integer;
begin

  if TOBCodeBarre.detail.count <> 0 then
    Grille.RowCount := TOBCodeBarre.detail.count + 1
  else
    InsertOnClick(Self);

  Grille.DoubleBuffered := true;
  Grille.BeginUpdate;

  TRY
    Grille.SynEnabled := false;
    for Indice := 0 to TOBCodeBarre.detail.count -1 do
    begin
      Grille.row := Indice+1;
      Grille.Cells [1,Grille.row] := TOBCodeBarre.Detail[Indice].GetString('BCB_QUALIFCODEBARRE');
      Grille.Cells [2,Grille.Row] := TOBCodeBarre.Detail[Indice].GetString('BCB_CODEBARRE');
    end;
  FINALLY
    Grille.SynEnabled := true;
    Grille.EndUpdate;
  END;

  TFVierge(ecran).HMTrad.ResizeGridColumns(Grille);

  //on Se positionne sur la première ligne de la grille
  Grille.row := 1;

  //
  ChargeGrilleToScreen(1);
  
end;

procedure TOF_BTCODEBARRE.GetObjects;
begin

  NatureCAB     := THEdit(GetControl('BCB_NATURECAB'));
  IdentifCAB    := THEdit(GetControl('BCB_IDENTIFCAB'));
  CodeBarre     := THEdit(GetControl('BCB_CODEBARRE'));
  //
  NomID         := THLabel(GetControl('NOMID'));
  //
  BValider      := TToolbarButton97 (GetControl('BVALIDER_'));
  BDelete       := TToolbarButton97 (GetControl('BDelete'));
  BInsert       := TToolbarButton97 (GetControl('BInsert'));
  //
  CabPrincipal  := THCheckbox(GetControl('BCB_CABPRINCIPAL'));
  //
  QualifCAB     := THValComboBox(GetControl('BCB_QUALIFCODEBARRE'));
  //
  Grille        := THGrid(GetControl('FListe'));
end;

Procedure TOF_BTCODEBARRE.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if Valeur='CREATION'          then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ = 'ID'   then
  Begin
    NatureCAB.Text  := Valeur;
    If      Valeur = 'FOU' Then Ecran.caption := 'Code Barre Fournisseurs'
    else If Valeur = 'CLI' Then Ecran.caption := 'Code Barre Clients'
    else if Valeur = 'MAR' Then Ecran.caption := 'Code Barre Marchandises'
    else if Valeur = 'CAT' Then Ecran.caption := 'Code Barre Catalogues Fournisseur'
    else if Valeur = 'AFF' Then Ecran.caption := 'Code Barre Chantiers'
    else                        Ecran.caption := 'Code Barre';
  end;

  if Champ = 'CODE' then IdentifCAB.Text := Valeur;

end;

procedure TOF_BTCODEBARRE.SetGrilleEvents (Etat : boolean);
begin

  if Etat then
  begin
    Grille.OnRowEnter := GrilleRowEnter;
  end
  else
  begin
    Grille.OnRowEnter := Nil;
  end;

end;

procedure TOF_BTCODEBARRE.GrilleRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin

  ChargeScreenToGrille;

  ChargeGrilleToScreen(Ou);

end;

procedure TOF_BTCODEBARRE.ChargeGrilleToScreen(Arow : Integer);
begin

  if Action = TaConsult then Exit;

  if Grille.RowCount = 0 then QualifCAB.SetFocus;

  //On charge une tob avec l'enregistrement approprié pour charger les zones écran
  TOBL := TOBCodeBarre.detail[Arow-1];

  QualifCAB.Value := TobL.GetString('BCB_QUALIFCODEBARRE');
  CodeBarre.text  := TobL.GetString('BCB_CODEBARRE');

  If TOBL.GetString('BCB_CABPRINCIPAL')='X' then CabPrincipal.Checked := True else CabPrincipal.Checked := False;

end;

procedure TOF_BTCODEBARRE.ChargeScreenToGrille;
begin

  //chargement sans distinction si création ou modification afin de simplifier les traitements
  TOBL.PutValue('BCB_NATURECAB',       NatureCAB.text);
  TOBL.PutValue('BCB_IDENTIFCAB',      IdentifCAB.text);
  if CabPrincipal.checked then TOBL.PutValue('BCB_CABPRINCIPAL',    'X') else TOBL.PutValue('BCB_CABPRINCIPAL',    '-');
  TOBL.PutValue('BCB_QUALIFCODEBARRE', QualifCAB.Value);
  TOBL.PutValue('BCB_CODEBARRE',       CodeBarre.text);

end;


procedure TOF_BTCODEBARRE.SetScreenEvents;
begin

  BValider.OnClick := ValideOnClick;

  BInsert.OnClick  := InsertOnClick;
  BDelete.OnClick  := DeleteOnClick;

  CodeBarre.OnExit := CodeBarreOnExit;

  CabPrincipal.OnClick := PrincipalOnClick;

end;

Procedure TOF_BTCODEBARRE.PrincipalOnClick(Sender : Tobject);
Var TOBLigne : TOB;
begin

  if Action = TaConsult then Exit;

  //nous ne pouvons avoir qu'un seul code barre principal
  if CabPrincipal.Checked then
  begin
    TOBLigne := TOBCodeBarre.FindFirst(['BCB_CABPRINCIPAL'],['X'],False);
    if TOBLigne <> nil then
    begin
      If TOBLigne.GetString('BCB_CODEBARRE') <> CodeBarre.Text then
      begin
        PGIError(TexteMessage[6], 'Code Barre Principal');
        CabPrincipal.Checked := False;
      End;
    end;
  end;

end;

procedure TOF_BTCODEBARRE.CodeBarreOnExit(Sender : Tobject);

begin

  If ControlCAB <> 0 then
  begin
    CodeBarre.text := '';
    CodeBarre.SetFocus;
  end;

end;

Function TOF_BTCODEBARRE.ControlCAB : Integer;
Var TypeCAB : string;
    Code    : string;
begin

  Result  :=0;

  TypeCAB := QualifCAB.Value;
  Code    := CodeBarre.Text;

  if (CodeBarre.text <> '') then
  begin
    if ((Code <> '') and (TypeCAB = '')) or ((Code = '') and (TypeCAB <> '')) then
    begin
      Result := 1;
      PGIError(TexteMessage[1], 'Contrôle Code Barre');
      Exit;
    end
    else
    begin
      Result := ControlCodeBarre(Code, TypeCAB);
      case Result of
        1 : BEGIN PGIError(TexteMessage[2], 'Contrôle Code Barre'); END;
        2 : BEGIN PGIError(TexteMessage[3], 'Contrôle Code Barre'); END;
        3 : BEGIN PGIError(TexteMessage[4], 'Contrôle Code Barre'); END;
        4 : BEGIN PGIError(TexteMessage[5], 'Contrôle Code Barre'); END;
      end;
      Exit;
    end;
  end;

end;


procedure TOF_BTCODEBARRE.OnCancel;
begin
  inherited;

end;

procedure TOF_BTCODEBARRE.OnClose;
begin
  inherited;

  DestroyTOB;

end;

procedure TOF_BTCODEBARRE.DeleteOnClick(Sender : Tobject);
begin
  inherited;

  if Action = taConsult then Exit;

  if TOBL = nil then exit;

  Ok_Action := 'Suppression';

  if PGIAsk(TexteMessage[6], 'Suppression')=MrYes then
  begin
    TOBL.DeleteDB;
    FreeandNil(TOBL);
  end;

  ChargementGrille;

end;

procedure TOF_BTCODEBARRE.InsertOnClick(Sender : Tobject);
begin

  if Action = taConsult then Exit;

  Ok_Action := 'Creation';

  TOBL := TOB.Create('BTCODEBARRE',TOBCodeBarre,-1);

  //Remise à zéro des zones écran...
  CabPrincipal.Checked := False;

  QualifCAB.Value := 'E13';
  CodeBarre.Text  := '';

  BValider.Visible:= True;

  Binsert.Visible := False;
  BDelete.Visible := False;

end;

procedure TOF_BTCODEBARRE.ValideOnClick(Sender : Tobject);
begin

  if Action = taConsult then Exit;

  If ControlCAB <> 0 then
  begin
    CodeBarre.text := '';
    CodeBarre.SetFocus;
  end
  Else
  begin
    ChargeScreenToGrille;

    TOBL.SetAllModifie(True);
    TOBL.InsertOrUpdateDB;

    ChargementGrille;

    BValider.Visible:= False;
    //
    Binsert.Visible := True;
    BDelete.Visible := True;
  end;

end;

Initialization
  registerclasses ( [ TOF_BTCODEBARRE ] ) ;
end.

