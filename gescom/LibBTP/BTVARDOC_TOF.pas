{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BVARIABLES (BVARIABLES)
Mots clefs ... : TOM;BVARIABLES
*****************************************************************}
Unit BTVARDOC_TOF  ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Vierge,
     HTB97,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     maineagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Grids,
     Graphics,
     Types,
     UtilsGrille,
     StrUtils,
     HSysMenu,
     HPanel,
     UTOF ;

Type
  TOF_BTVARDOC  = Class (TOF)

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
    TypeVariable  : String;
    CodeArt       : String;
    //
    IndexFile     : Integer;
    //
    GrilleVar     : TGestionGS;
    //
    TobVardoc     : TOB;
    //
    // Les Boutons d'Action
    BInsert       : TToolbarButton97;
    BDelete       : TToolbarButton97;
    BValider      : TToolbarButton97;
    BFermer       : TToolbarButton97;
    //
    FListe        : THGrid;
    //
    NaturePiece   : THEdit;
    Souche        : THEdit;
    Numero        : THEdit;
    Indice        : THEdit;
    Numordre      : THEdit;
    UniqueBlo     : THEdit;
    //
    CodeVariable  : THEdit;
    Libelle       : THEdit;
    Valeur        : THEdit;
    //
    Cancel        : Boolean;
    //
    //Action        : TActionFiche;
    //
    procedure ChargementEcran;
    function DecoupeCritere(var critere: string): string;
    procedure DeleteOnClick(Sender: Tobject);
    procedure FermerOnClick(Sender: Tobject);
    procedure GetObjects;
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure InitGrillevar;
    procedure InsertOnClick(Sender: TObject);
    procedure SetGrilleEvents(Etat: Boolean);
    procedure SetScreenEvents;
    procedure ValiderOnClick(sender: TObject);                             
    //

    //

    end ;

Implementation

const SG_CODE     = 1;
      SG_LIBELLE  = 2;
      SG_VALEUR   = 3;


procedure TOF_BTVARDOC.OnArgument ( S: String ) ;
var Critere  : String;
begin
  Inherited ;

  Cancel := False;

  GetObjects;

  SetScreenEvents;

  SetGrilleEvents(True);

  // Chargement du critére de sélection
  Critere:=uppercase(Trim(ReadTokenSt(S))) ;
  TypeVariable := Critere;

  Critere:= uppercase(Trim(ReadTokenSt(S)));
  CodeArt:= Critere;

  if TypeVariable = 'D' then
    TFVierge(Ecran).Caption := 'Variables document : ' + CodeArt
  else if TypeVariable = 'LS' then
    TFVierge(Ecran).Caption := 'Variables Lignes sous-Détail'
  else
    TFVierge(Ecran).Caption := 'Variables Lignes';


  //on redécoupe le numéro de dossier passé
  if Critere <> '' then  NaturePiece.Text := decoupeCritere(Critere);
  if Critere <> '' then  Souche.Text      := decoupeCritere(Critere);
  if Critere <> '' then  Numero.Text      := decoupeCritere(Critere);
  if Critere <> '' then  Indice.Text      := decoupeCritere(Critere);
  if Critere <> '' then  Numordre.Text    := decoupeCritere(Critere);
  if Critere <> '' then  UniqueBlo.Text   := decoupeCritere(Critere);

  TOBVardoc := LaTob;

  if TobVardoc.Detail.count <> 0 then
    IndexFile := TobVardoc.Detail[TobVardoc.Detail.Count-1].GetValue('BVD_INDEXFILE')
  else
    IndexFile := 0;

  //Initialisation des grilles
  InitGrilleVar;

  //chargement des zones écran avec les zones tob
  ChargementEcran;

  if TobVarDoc.Detail.count > 0 then
    GSRowEnter(self, 1, Cancel, false)
  else
    InsertOnClick(Self);

end ;

Procedure TOF_BTVARDOC.GetObjects;
begin

    Bdelete       := Ttoolbarbutton97(GetControl('BDELETE'));
    BValider      := TtoolBarButton97(GetControl('BVALIDER'));
    BInsert       := TToolBarButton97(GetControl('BINSERT'));
    BFermer       := TToolBarButton97(GetControl('BFERME'));
    //
    FListe        := THGrid(GetControl('GRILLE_VAR'));
    //
    CodeVariable  := THEdit(GetControl('BVD_CODEVARIABLE'));
    Libelle       := THEdit(GetControl('BVD_LIBELLE'));
    Valeur        := THEdit(GetControl('BVD_VALEUR'));
    NaturePiece   := THEdit(GetControl('BVD_NATUREPIECE'));
    Souche        := THEdit(GetControl('BVD_SOUCHE'));
    Numero        := THEdit(GetControl('BVD_NUMERO'));
    Indice        := THEdit(GetControl('BVD_INDICE'));
    NumOrdre      := THEdit(GetControl('BVD_NUMORDRE'));
    UniqueBlo     := THEdit(GetControl('BVD_UNIQUEBLO'));

end;

Procedure TOF_BTVARDOC.SetScreenEvents;
begin

  BInsert.OnClick   := InsertOnClick;
  BValider.OnClick  := ValiderOnClick;
  BFermer.OnClick   := FermerOnClick;
  bdelete.OnClick   := DeleteOnClick;

end;

Procedure TOF_BTVARDOC.SetGrilleEvents(Etat : Boolean);
begin

  if Etat then
  begin
    FListe.OnRowEnter      := GSRowEnter;
  end
  else
  begin
    FListe.OnRowEnter      := Nil;
  end;

end;


Procedure TOF_BTVARDOC.InitGrillevar;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GrilleVar                := TGestionGS.Create;

  GrilleVar.Ecran          := TFVierge(Ecran);
  GrilleVar.GS             := Fliste;
  GrilleVar.TOBG           := TobVardoc;

  // Définition de la liste de saisie pour la grille Détail
  GrilleVar.ColNamesGS     := 'BVD_CODEVARIABLE;BVD_LIBELLE';
  GrilleVar.alignementGS   := 'C.0  ---;G.0  ---;';
  GrilleVar.LibColNameGS   := 'Code;Désignation;';
  GrilleVar.LargeurGS      := '50;100;';
  //
  GrilleVar.RowHeight := 18;

  GrilleVar.DessineGrille;

  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(FListe);

  FListe.FixedRows := 1;

end;

procedure TOF_BTVARDOC.ChargementEcran;
var ind     : integer;
begin

  With GrilleVar do
  begin
    if TOBG.Detail.count <> 0 then
      GS.RowCount := TOBG.detail.count + 1
    else
      Exit;
    //
    GS.DoubleBuffered := true;
    GS.BeginUpdate;
    //
    TRY
      GS.SynEnabled := false;
      for Ind := 0 to TOBG.detail.count -1 do
      begin
        GS.row := Ind+1;
        TOBG.PutGridDetail(GS,False,True,ColNamesGS);
      end;
    FINALLY
      GS.SynEnabled := true;
      GS.EndUpdate;
    END;
    TFVierge(Tform).HMTrad.ResizeGridColumns(GS);
    GS.Row := 1;
  end;

end;

procedure TOF_BTVARDOC.OnClose ;
begin
  Inherited ;

end ;

procedure TOF_BTVARDOC.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

  with GrilleVar do
  begin
    TOBG.detail[GS.row-1].PutEcran(Ecran, THPanel(GetControl('Panel1')));
    if TOBG.detail[GS.row-1].GetValue('BVD_QTEEXPPOUR') = 'X' then
      setcontrolvisible('BVD_QTEEXPPOUR', true)
    else
      setcontrolvisible('BVD_QTEEXPPOUR', False);
  end;

end;

procedure TOF_BTVARDOC.InsertOnClick(Sender: TObject);
begin

  BInsert.visible := false;
  BDelete.Visible := False;

  CodeVariable.Text := '';
  Libelle.Text      := '';
  Valeur.Text       := '';
  //
  SetControlText('BVD_QTEEXPPOUR', '');

  CodeVariable.SetFocus;

end;

Procedure Tof_BTVARDOC.ValiderOnClick(sender : TObject);
var TobFille : tob;
begin

  With GrilleVar do
  begin
    if Binsert.Visible then
      TOBG.Detail[Gs.row - 1].GetEcran(Ecran, THPanel(GetControl('Panel1')))
    Else
    begin
      GS.InsertRow(TobVarDoc.Detail.Count + 1);
      TobFille := Tob.Create('BVARDOC', TobVarDoc, -1);
      Inc(IndexFile);
      TobFille.Putvalue('BVD_NATUREPIECE',  NaturePiece.text);
      TobFille.Putvalue('BVD_SOUCHE',       Souche.text);
      TobFille.Putvalue('BVD_NUMERO',       Numero.text);
      TobFille.Putvalue('BVD_INDICE',       Indice.text);
      TobFille.Putvalue('BVD_NUMORDRE',     Numordre.text);
      TobFille.Putvalue('BVD_UNIQUEBLO',    Uniqueblo.text);
      TobFille.Putvalue('BVD_ARTICLE',      CodeArt);
      TobFille.Putvalue('BVD_INDEXFILE',    IndexFile);
      TobFille.Putvalue('BVD_CODEVARIABLE', CodeVariable.text);
      TobFille.Putvalue('BVD_LIBELLE',      Libelle.text);
      TobFille.Putvalue('BVD_VALEUR',       Valeur.Text);
      Tobfille.SetAllModifie(true);
      TobFille.InsertOrUpdateDB(true);
    end;

    ChargementEcran;

    BInsert.Visible := True;
    BDelete.Visible := True;

    if TobVardoc.Detail.Count = 0 then
      InsertOnClick(self)
    else
      GSRowEnter(self, 1, Cancel, false);
  end;

end;

Procedure TOF_BTVARDOC.FermerOnClick(Sender : Tobject);
begin

  TheTOB := TOBVarDoc;

end;

Procedure TOF_BTVARDOC.DeleteOnClick(Sender : Tobject);
Var TextMsg : string;
begin

  with GrilleVar do
  begin
    if Tobg.Detail.Count = 0 then exit;
    //
    if TypeVariable = 'D' then
      TextMsg := 'Confirmez-vous la suppression de la variable document ?'
    else if TypeVariable = 'LS' then
      TextMsg := 'Confirmez-vous la suppression de la variable Ligne sous-Détail ?'
    else
      TextMsg := 'Confirmez-vous la suppression de la variable Ligne ?';
    //
    if PGIask(TraduireMemoire(TextMsg),'Suppression Variables') = mrNo then exit;
    //
    TOBG.Detail[Gs.Row-1].free;
    //
    GS.DeleteRow(Gs.Row);
    //
    ChargementEcran;
    //
    if TOBG.Detail.Count = 0 then
      InsertOnClick(Self)
    else
      GSRowEnter (self, 1, Cancel, false);
  end;

end;

procedure TOF_BTVARDOC.OnCancel;
begin
  inherited;
end;

procedure TOF_BTVARDOC.OnDelete;
begin
  inherited;

end;

procedure TOF_BTVARDOC.OnDisplay;
begin
  inherited;

end;

procedure TOF_BTVARDOC.OnLoad;
begin
  inherited;

end;

procedure TOF_BTVARDOC.OnNew;
begin
  inherited;

end;

procedure TOF_BTVARDOC.OnUpdate;
begin
  inherited;

end;

Function Tof_BTVARDOC.DecoupeCritere(Var critere : string) : string;
var ChampMul : String;
    x       : Integer;
begin

  result := '';

  x := pos('-', Critere);

  if x <> 0 then
  begin
    ChampMul := copy(Critere, 1, x - 1);
    Result   := ChampMul;
    Critere  := copy(Critere, x + 1, length(Critere));
  end
  else Result := Critere;

end;

Initialization
  registerclasses ( [ TOF_BTVARDOC ] ) ;
end.

