{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMODELHEBDO ()
Mots clefs ... : TOF;BTMODELHEBDO
*****************************************************************}
Unit BTMODELHEBDO_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     Grids,
     Windows,
     Messages,
{$IFNDEF EAGLCLIENT}
     db,
     fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
     MainEagl,
{$ENDIF}
     forms,
     LookUp,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UtilModeleConso,
     UTOB,
     vierge,
     menus,
     UTOF ;

Type

  TModeAction = (TmaCreat,TmaModif);
  TOF_BTMODELHEBDO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    // Controles de la fiche
    POPCHOIX : TpopupMenu;
    BMS_TYPERESSOURCE : THValCombobox;
    BMS_CODEMODELE,BMS_LIBELLE : Thedit;
    GS : THgrid;
    BInsert,Bdelete,BCherche,Bannul,Bvalider,BFerme : TToolbarButton97;
    MnPrest,MnFrais : TmenuItem;
    //
    stPrev : string;
    TOBModele : TOB;
    nbcolsInListe : integer;
    ModeAction : TModeAction;
    NomList, FNomTable, FLien, FSortBy, stCols, stColListe : string;
    Ftitre : Hstring;
    FLargeur, FAlignement, FParams, FPerso: string;
    title, NC : Hstring;
    OkTri, OkNumCol : boolean;
    //
    G_NUMORDRE,G_CODEARTICLE,G_LIBELLE,G_QTE1,G_QTE2,G_QTE3,G_QTE4,G_QTE5,G_QTE6,G_QTE7 : integer;
    //
    function ValidelaCreation : boolean;
    procedure SetMode (TheModeAction : TModeAction);
    procedure ChargeLaGrille;
    procedure CodeModeleExit (Sender : TObject);
    procedure GetComponents;
    procedure SetEvents;
    procedure TypeRessourceExit (Sender : Tobject);
    procedure BinsertClick (Sender : TObject);
    procedure LibelleModeleExit (Sender : Tobject);
    procedure SetModeSaisieGrid (Etat : boolean);
    procedure SetparamGrille;
    procedure DefiniGrille;
    procedure initgrille;
    procedure AfficheLigne(Grid: THgrid; TOBL: TOB; indice: integer);
    procedure Remplitgrille;
    function AjouteLigne: TOB;
    procedure EventGrille(Active: boolean);
    procedure PositionneDansGrid(TheGrid: THGrid; Arow, Acol: integer);
    procedure BChercheClick (Sender : TOBject);
    procedure BAnnuleClick (Sender : TOBject);
    //
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
    procedure GSElipsisClick(Sender: TOBject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
		procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    //
    function IsLigneVide (Arow : integer) : boolean;
    function ControleLigne(ligne: integer; var Acol: integer): boolean;
    function ZoneAccessible(Grille: THgrid; var ACol,
      ARow: integer): boolean;
    procedure ZoneSuivanteOuOk(Grille: THGrid; var ACol, ARow: integer;
      var Cancel: boolean);
    function FindCodeArticle(TOBL: TOB; Valeur: string): boolean;
    procedure RechFRais(Sender: TObject);
    procedure RechPrestations(Sender: TObject);
    procedure DeleteLigne;
    procedure reindiceGrid;
    procedure AfficheLagrille;
    procedure BeforeValide;
    procedure ValideLeModele;
    procedure detruitAncienModele;
    procedure EcritModele;
    procedure NextModele;
    procedure GSEnter(Sender: TObject);
    procedure SetButton(Entete: boolean);
    function ControleEntete: boolean;
  end ;
Implementation

procedure TOF_BTMODELHEBDO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMODELHEBDO.OnDelete ;
begin
  Inherited ;
  //
  DetruitAncienModele;
  NextModele;
end ;

procedure TOF_BTMODELHEBDO.OnUpdate ;
var Indice,Acol : integer;
begin
  Inherited ;
  for Indice := 0 to tobModele.detail.count-1 do
  begin
    if not ControleLigne (Indice+1,Acol) then
    begin
      TFVierge(Ecran).ModalResult := 0;
      PositionneDansGrid (GS,Indice+1,Acol);
      exit;
    end;
  end;

  BeforeValide;

  if TRANSACTIONS (ValideLeModele,0) <> OeOk then BEGIN PgiBox('Erreur lors de l''ecriture'); END;
  NextModele;
  TFVierge(Ecran).ModalResult := 0;
end ;

procedure TOF_BTMODELHEBDO.OnLoad ;
begin
  Inherited ;
  SetModeSaisieGrid (false);
  SetEvents;
  SetMode(TmaModif); // par defaut
end ;

procedure TOF_BTMODELHEBDO.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBModele := TOB.Create ('BTMODELSHEBDO',nil,-1);
  GetComponents;
  SetButton (True);

  SetparamGrille;
end ;

procedure TOF_BTMODELHEBDO.OnClose ;
begin
  if TOBmodele.detail.count > 0 then
  begin
    if PgiAsk (TraduireMemoire('Une saisie est en cours. Abandonner ?'))<>MrYes then
    begin
      TFVierge(Ecran).ModalResult := 0;
      exit;
    end;
  end;
  TOBModele.free;
  Inherited ;
end ;

procedure TOF_BTMODELHEBDO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMODELHEBDO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMODELHEBDO.GetComponents;
begin
  BMS_TYPERESSOURCE := THValComboBox(GetCOntrol('BMS_TYPERESSOURCE'));
  BMS_CODEMODELE := THEdit(GetCOntrol('BMS_CODEMODELE'));
  BMS_LIBELLE := THEdit(GetCOntrol('BMS_LIBELLE'));
  BInsert := TToolbarButton97 (GetControl('BInsert'));
  BDelete := TToolbarButton97 (GetControl('BDElete'));
  BCherche := TToolbarButton97 (GetControl('BCherche'));
  BAnnul := TToolbarButton97 (GetControl('BAnnul'));
  Bvalider := TToolbarButton97 (GetControl('Bvalider'));
  BFerme := TToolbarButton97 (GetControl('BFerme'));
  POPCHOIX := TPopupMenu (GetCOntrol('POPCHOIX'));
  GS := THGrid(GetControl('GS'));
  MnPrest := TMenuItem(GetControl('MNPREST'));
  MnFrais := TMenuItem(GetControl('MNFRAIS'));

end;

procedure TOF_BTMODELHEBDO.SetEvents;
begin
  BMS_TYPERESSOURCE.OnChange := TypeRessourceExit;
  BMS_CODEMODELE.OnExit := CodeModeleExit;
  BMS_LIBELLE.OnExit := LibelleModeleExit;
  BInsert.onClick := BinsertClick;
  MnPrest.onClick := RechPrestations;
  MnFrais.onClick := RechFrais;
  Bcherche.onclick := BChercheClick;
  Bannul.onclick := BAnnuleClick;
  TFVierge(Ecran).OnKeyDown := FormKeyDOwn;
end;

procedure TOF_BTMODELHEBDO.TypeRessourceExit(Sender: Tobject);
begin
  BMS_CODEMODELE.Text := '';
  BMS_LIBELLE.Text := '';
  BMS_CODEMODELE.Plus := ' AND BMS_TYPERESSOURCE="'+BMS_TYPERESSOURCE.Value+'"';
end;

procedure TOF_BTMODELHEBDO.CodeModeleExit(Sender: TObject);
begin
  if (ModeAction=TmaModif) then
  begin
    if BMS_CODEMODELE.text = '' then exit;
    if  (Not IsModeleExist (BMS_CODEMODELE.Text,BMS_TYPERESSOURCE.Value)) then
    begin
      PgiBox(TraduireMemoire('Ce modèle n''existe pas'));
      BMS_CODEMODELE.Text := '';
      SetFocusControl('BMS_CODEMODELE');
      exit;
    end;
		BMS_LIBELLE.text := GetLibelleModele (BMS_CODEMODELE.Text,BMS_TYPERESSOURCE.Value);
  end else
  begin
    // Creation
    if (IsModeleExist (BMS_CODEMODELE.Text,BMS_TYPERESSOURCE.Value)) then
    begin
      PgiBox(TraduireMemoire('Ce modèle existe déjà.'));
      BMS_CODEMODELE.Text := '';
      SetFocusControl('BMS_CODEMODELE');
      exit;
    end;
  end;
end;

procedure TOF_BTMODELHEBDO.ChargeLaGrille;
var Arow,Acol : integer;
begin
  DefiniGrille;
  InitGrille;
  //
  Remplitgrille;
  SetModeSaisieGrid (true);
  EventGrille(true);
  Arow := 1;
  Acol := 1;
  PositionneDansGrid (GS,Arow,Acol);
end;

procedure TOF_BTMODELHEBDO.BinsertClick(Sender: TObject);
begin
  //
  SetMode (TmaCreat);
  BMS_CODEMODELE.Text := '';
end;

procedure TOF_BTMODELHEBDO.SetMode(TheModeAction : TmodeAction);
begin
  ModeAction := TheModeAction;
  BMS_LIBELLE.Enabled := (ModeAction=TmaCreat);
  BMS_CODEMODELE.ElipsisButton := (ModeAction=TmaModif);
  BInsert.Enabled := (ModeAction<>TmaCreat);

end;

procedure TOF_BTMODELHEBDO.LibelleModeleExit(Sender: Tobject);
begin
end;

procedure TOF_BTMODELHEBDO.SetModeSaisieGrid(Etat: boolean);
begin
  GS.visible := (Etat);
  GS.Enabled := Etat;
  BMS_TYPERESSOURCE.Enabled := not Etat;
  BMS_CODEMODELE.Enabled := not Etat;
  BMS_LIBELLE.Enabled := not Etat;
  BInsert.Enabled := not Etat;
  Bdelete.Enabled := (Etat) and (ModeAction<>TmaCreat);
end;


procedure TOF_BTMODELHEBDO.SetparamGrille;
var LaListe,lelement : string;
begin
  // récupération du paramétrage général des grilles
  NomList := 'BTMODELEHEBDO';
  ChargeHListe (NomList, FNomTable, FLien, FSortBy, stCols, FTitre,
                FLargeur, FAlignement, FParams, title, NC, FPerso, OkTri, OkNumCol);
  stColListe := stCols;
  nbcolsInListe := 0;
  LaListe := stColListe;
  repeat
    lelement := READTOKENST (laListe);
    if lelement <> '' then
    begin
      if lelement = 'BSM_NUMORDRE' then G_NUMORDRE := nbcolsinListe else
      if lelement = 'BSM_CODEARTICLE' then G_CODEARTICLE := nbcolsinListe else
      if lelement = 'BSM_LIBELLE' then G_LIBELLE := nbcolsinListe else
      if lelement = 'BSM_QTEJ1' then  G_QTE1 := nbcolsinListe else
      if lelement = 'BSM_QTEJ2' then  G_QTE2 := nbcolsinListe else
      if lelement = 'BSM_QTEJ3' then  G_QTE3 := nbcolsinListe else
      if lelement = 'BSM_QTEJ4' then  G_QTE4 := nbcolsinListe else
      if lelement = 'BSM_QTEJ5' then  G_QTE5 := nbcolsinListe else
      if lelement = 'BSM_QTEJ6' then  G_QTE6 := nbcolsinListe else
      if lelement = 'BSM_QTEJ7' then  G_QTE7 := nbcolsinListe ;
      inc(nbcolsInListe);
    end;
  until lelement = '';
end;


procedure TOF_BTMODELHEBDO.DefiniGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec : integer;
    indice : integer;
    FFQTE,FFLQTE : string;
begin
  FFQTE := '###';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQTE := FFQTE+'0.';
    for indice := 1 to V_PGI.OkDecQ do
    begin
      FFQTE := FFQTE + '0';
    end;
  end;

  GS.ColCount := NbColsInListe;

  st := stColListe;
  lesalignements := Falignement;
  lestitres := FTitre;
  leslargeurs := flargeur;

  for indice := 0 to nbcolsInListe -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;

    GS.cells[Indice,0] := leTitre;

    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;

    GS.ColWidths[indice] := strtoint(lalargeur);

    if OkLib then GS.ColFormats[indice] := 'CB=' + Get_Join(Nam)
    else if (Dec<>0) or (Sep) then GS.ColFormats[indice] := FF ;

    if Copy(nam,1,7) = 'BSM_QTE' then
    begin
      if OkNulle then FFLQte := FFQTE+';-'+FFQTE+'; ;';
      GS.ColFormats[indice] := FFLQTE ;
    end;
  end ;
end;


procedure TOF_BTMODELHEBDO.initgrille;
begin
  GS.VidePile (false);
end;

function TOF_BTMODELHEBDO.AjouteLigne : TOB;
begin
  result := TOB.create ('BTDETMODELSHEB',TOBModele,-1);
  result.putValue('BSM_NUMORDRE',TOBModele.detail.count);
end;


procedure TOF_BTMODELHEBDO.Remplitgrille;
begin
  GS.VidePile (false);
  if TOBModele.detail.count = 0 then AjouteLigne;
  if TOBModele.detail.count > 0 then GS.RowCount := TOBModele.detail.count +2;
  if GS.RowCount < 2 then GS.rowcount := 3;
  AfficheLagrille;
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
end;

procedure TOF_BTMODELHEBDO.AfficheLigne(Grid : THgrid;TOBL : TOB;indice : integer);
begin
  TOBL.PutLigneGrid (Grid,Indice+1,false,false,stColListe);
end;


function TOF_BTMODELHEBDO.ValidelaCreation: boolean;
begin
  result := (PgiAsk ('Vous êtes sur le point de créer un modèle. Valider ?')=MrYes);
end;


procedure TOF_BTMODELHEBDO.EventGrille(Active : boolean);
begin
  if Active then
  begin
    GS.OnEnter  := GSEnter;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.OnCellEnter  := GSCellEnter;
    GS.OnCellExit  := GSCellExit;
    GS.OnElipsisClick := GSElipsisClick;
    GS.PostDrawCell := GSPostDrawCell;
    GS.OnKeyDown := GSKeyDOwn;
  end
  else
  begin
    GS.OnEnter  := nil;
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnCellEnter  := nil;
    GS.OnCellExit  := nil;
    GS.OnElipsisClick := nil;
    GS.PostDrawCell := nil;
  end;
end;


procedure TOF_BTMODELHEBDO.PositionneDansGrid (TheGrid : THGrid;Arow,Acol : integer);
var cancel : boolean;
begin
  TheGrid.CacheEdit;
  EventGrille (false);
  TheGrid.row := Arow;
  TheGrid.col := Acol;
  EventGrille (True);
  TheGrid.OnRowEnter (self,Arow,Cancel,false);
  TheGrid.OnCellEnter (self,Acol,Arow,Cancel);
  TheGrid.row := Arow;
  TheGrid.col := Acol;
  TheGrid.ShowEditor;
end;

procedure TOF_BTMODELHEBDO.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
    OkElipsis : boolean;
begin
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  if not Cancel then
  begin
    TOBL := TOBModele.detail[Arow-1];
    if TOBL <> nil then
    begin

      if (TOBL.GetValue('BSM_ARTICLE')='') and (Acol > G_CODEARTICLE) then BEGIN Acol := G_CODEARTICLE; PositionneDansGrid (GS,Arow,Acol); END;
      if (TOBL.GetValue('BSM_ARTICLE')<>'') And (Acol = G_CODEARTICLE) then BEGIN Acol := G_LIBELLE; PositionneDansGrid (GS,Arow,Acol); END;
    end;
    GS.ElipsisButton := (GS.col = G_CODEARTICLE);
    stprev := GS.Cells [Acol,Arow];
  end;
end;

procedure TOF_BTMODELHEBDO.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
Var TOBLigne : TOB;
    indice : integer;
begin

//  if Action = taConsult then Exit;

  // pas de controle si remonté et ligne vide
  if (GS.row < Arow) and (IsLigneVide (Arow))  then exit;

  if (stPrev <> GS.cells[Acol,Arow]) or (GS.cells[Acol,Arow]= '') then
  begin
    TOBLigne := TOBModele.detail[Arow-1];
    if ACol = G_CODEARTICLE then
    Begin
      if not FindCodeArticle (TOBLigne,GS.cells[Acol,Arow]) then cancel := true;
    end;
    if (Acol >= G_QTE1) and (Acol <= G_QTE7) then
    begin
      Indice := (Acol-G_QTE1)+1;
      TOBLIGNE.putValue('BSM_QTEJ'+IntToStr(Indice),valeur(GS.cells[Acol,Arow]));
    end;
    if (Acol = G_LIBELLE) then
    begin
      TOBLIGNE.putValue('BSM_LIBELLE',GS.cells[Acol,Arow]);
    end;
    AfficheLigne (GS,TOBLIGNE,Arow-1);
    if not cancel then
    begin
      stPrev := GS.cells[Acol,Arow];
    end;
  end;
end;

procedure TOF_BTMODELHEBDO.GSElipsisClick(Sender: TOBject);
var Coords : Trect;
    PointD,PointF : Tpoint;
begin
  if GS.col=G_CODEARTICLE then
  begin
    if (Pos(BMS_TYPERESSOURCE.value,'SAL;INT')= 0) then
    begin
      RechPrestations (self);
      exit;
    end;

    Coords := GS.CellRect (GS.col,GS.row);
    PointD := GS.ClientToScreen ( Coords.Topleft)  ;
    PointF := GS.ClientToScreen ( Coords.BottomRight )  ;
    POPCHOIX.Popup (pointF.X ,pointD.y+10);
  end;
end;

procedure TOF_BTMODELHEBDO.GSPostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
begin
end;

procedure TOF_BTMODELHEBDO.GSEnter(Sender: TObject);
var Acol,Arow : integer;
    cancel : boolean;
begin
  cancel := false;
  GS.row := 1;
  GS.col := 1;
  Arow := 1;
  Acol := 1;
  GS.cacheEdit;
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  cancel := false;
//  GSRowEnter (self,Arow,cancel,false);
//  GSCellEnter (self,Acol,Arow,cancel);
  GS.row := Arow;
  GS.col := Acol;
  GS.ElipsisButton := (Acol = G_CODEARTICLE)  ;
  stprev := GS.Cells [Acol,Arow];
  GS.ShowEditor;
  //
end;

procedure TOF_BTMODELHEBDO.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
Var TobLigne : Tob;
begin
  if (Ou > 1) and (IsLigneVide (Ou-1)) then
  BEGIN
    cancel := true;
    Exit;
  END;
  if (Ou >= GS.rowCount -1) then GS.rowCount := GS.rowCount +1;
  if Ou > TOBModele.Detail.count then
  begin
    TOBLigne := AjouteLigne;
    AfficheLigne (GS,TOBLigne,Ou-1);
  end;
end;

procedure TOF_BTMODELHEBDO.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var Acol : integer;
Begin
  if IsLigneVide (Ou) then exit;
  if not ControleLigne (Ou,Acol) then
  BEGIN
  	cancel := true;
    PositionneDansGrid (GS,Ou,Acol);
  END;
end;

function TOF_BTMODELHEBDO.IsLigneVide(Arow: integer): boolean;
var TOBL : TOB;
begin
  result := true;
  TOBL := TOBModele.detail[Arow-1];
  if TOBL.GetValue('BSM_CODEARTICLE') <> '' then result:= false;
end;

function TOF_BTMODELHEBDO.ControleLigne (ligne : integer; var Acol : integer) : boolean;
var TOBL : TOB;
    Indice : integer;
    CumulSem : double;
begin
  result := true;
  if IsLIgneVide (Ligne) then exit; // si elle est vide bon ..
  //
  TOBL := TOBModele.detail[Ligne-1];
  for Indice := 1 to 7 do
  begin
    CumulSem := CumulSem + TOBL.GetValue('BSM_QTEJ'+IntToStr(Indice));
  end;
  if CumulSem = 0 then
  begin
    PgiBox ('Vous devez renseigner des quantités');
    ACol := G_QTE1;
    result := false;
    exit;
  end;

end;

procedure TOF_BTMODELHEBDO.ZoneSuivanteOuOk(Grille : THGrid;var ACol, ARow : integer; var Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  ChgLig := (Grille.Row <> ARow);
  ChgSens := False;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol <= Grille.Col)) then Sens := 1;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  while not ZoneAccessible(Grille,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := Grille.RowCount ;
      // ---
      if ((ACol = Grille.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          // Ajout d'une ligne
          break;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.ColCount;
        ChgLig := False;
      end;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

//
function TOF_BTMODELHEBDO.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
var TOBL : TOB;
begin
  TOBL := nil;
  if Arow < TOBModele.detail.count then TOBL := TOBModele.detail[Arow-1];
  result := true;
  if Grille.ColWidths[acol] = 0 then BEGIN result := false; exit; END;
  if (Acol=G_CODEARTICLE) then
  BEGIN
    if (TOBL <> nil) and (TOBL.GetValue('BSM_ARTICLE')<>'') then
    begin
      result := false;
      exit;
    end;
  END;
  if (Acol>G_CODEARTICLE) then
  begin
    if (TOBL<>nil) and (TOBL.GetValue('BSM_ARTICLE')='') then
    begin
      result := false;
      exit;
    end;
  end;
end;

function TOF_BTMODELHEBDO.FindCodeArticle (TOBL : TOB; Valeur : string ) : boolean;
var QQ : TQuery;
    Req : STring;
begin
  result := true;
  if Valeur = '' then BEGIN result := false; Exit; END;
  Req := 'SELECT GA_CODEARTICLE,GA_TYPEARTICLE,GA_ARTICLE,GA_LIBELLE,BNP_TYPERESSOURCE,BNP_LIBELLE '+
         'FROM ARTICLE '+
         'LEFT JOIN NATUREPREST ON GA_NATUREPRES=BNP_NATUREPRES WHERE '+
         'GA_CODEARTICLE="'+Valeur+'" AND GA_STATUTART="UNI" '+
         'AND ((GA_TYPEARTICLE="PRE" AND BNP_TYPERESSOURCE="'+BMS_TYPERESSOURCE.Value+'") OR (GA_TYPEARTICLE="FRA"))';

  QQ := OpenSql (Req,True);
  result := not QQ.eof;
  if result then
  begin
    // verification de l'unicité de la ligne
    if TOBModele.findfirst(['BSM_ARTICLE'],[QQ.findField('GA_ARTICLE').asString],true) <> nil then
    begin
      PgiBox ('Cette prestation/Frais est déjà présente dans ce modèle');
      result := false;
    end;
  end;
  if Result then
  begin
    TOBL.PutValue('BSM_LIBELLE',QQ.findField('GA_LIBELLE').asString);
    TOBL.PutValue('BSM_ARTICLE',QQ.findField('GA_ARTICLE').asString);
    TOBL.PutValue('BSM_CODEARTICLE',QQ.findField('GA_CODEARTICLE').asString);
  end else
  begin
    TOBL.InitValeurs;
    TOBL.putValue('BSM_NUMORDRE',GS.row);
  end;
  ferme (QQ);
end;


Procedure TOF_BTMODELHEBDO.RechPrestations (Sender : TObject);
Var  stChamps : String;
     Article	: string;
     TOBL : TOB;
begin
  TOBL := TOBModele.detail[GS.Row-1];
  if GS.Cells [GS.Col,GS.row] <> '' then stchamps := 'GA_CODEARTICLE='+trim(GS.Cells[GS.col,GS.row])+';' else stchamps := '';
  stChamps := stChamps+'TYPERESSOURCE='+ BMS_TYPERESSOURCE.Value  ;
  stChamps := stChamps+';GA_TYPEARTICLE=PRE';
  //
  Article := AGLLanceFiche('BTP', 'BTPREST_RECH', '','',stChamps);
  //
  if Article <> '' then
  begin
    GS.cells[GS.col,GS.row] := Trim(Copy(Article,1,18));
    TOBL := TOBModele.detail[GS.row-1];
    if FindCodeArticle (TOBL,Trim(Copy(Article,1,18))) then
    begin
      AfficheLigne (GS,TOBL,GS.row-1);
      PositionneDansGrid (GS,GS.row,G_LIBELLE);
      stPrev := GS.cells[GS.col,GS.row];
    end;
  end;
end;

Procedure TOF_BTMODELHEBDO.RechFRais (Sender : TObject);
Var  stChamps : String;
     Article	: string;
     TOBL : TOB;
begin
  if GS.Cells [GS.Col,GS.row] <> '' then stchamps := 'GA_CODEARTICLE='+trim(GS.Cells[GS.col,GS.row])+';' else stchamps := '';

  stChamps := stchamps+'XX_WHERE=AND (GA_TYPEARTICLE="FRA")';
  Article := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps);
  //
  if Article <> '' then
  begin
    GS.cells[GS.col,GS.row] := Trim(Copy(Article,1,18));
    TOBL := TOBModele.detail[GS.row-1];
    if FindCodeArticle (TOBL,Trim(Copy(Article,1,18))) then
    begin
      AfficheLigne (GS,TOBL,GS.row-1);
      PositionneDansGrid (GS,GS.row,G_LIBELLE);
      stPrev := GS.cells[GS.col,GS.row];
    end;
  end;
end;

procedure TOF_BTMODELHEBDO.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    case Key of
    VK_DELETE: if (Shift = [ssCtrl]) then
    begin
      DeleteLigne;
      Key := 0;
    end;
    VK_F5: if (Shift = []) then
    begin
      if TFVierge(Ecran).ActiveControl = GS then
      begin
        GSElipsisClick(Sender);
        Key := 0;
      end;
    end;
    VK_ESCAPE:
    BEGIN
    	Key := 0;
//      BannulClick (self);
    END;
  end;

end;

procedure TOF_BTMODELHEBDO.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    case Key of
    VK_F5: if (Shift = []) then
    begin
      if TFVierge(Ecran).ActiveControl = GS then
      begin
        GSElipsisClick(Sender);
        Key := 0;
      end;
    end;
    VK_RETURN : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;

end;

procedure TOF_BTMODELHEBDO.DeleteLigne;
var LastRow : integer;
begin
  if TFVierge(Ecran).ActiveControl = GS then
  begin
    GS.DeleteRow (GS.row);
    if GS.row <= TOBModele.detail.count then
    begin
      TOBModele.detail[GS.row-1].free;
    end;
    ReindiceGrid;
    Remplitgrille;
    if lastrow > TOBModele.detail.count then lastRow := TOBModele.detail.count;
    PositionneDansGrid (GS,LastRow,1); 
  end;
end;

procedure TOF_BTMODELHEBDO.ReindiceGrid;
var Indice : integer;
    TOBL : TOB;
begin
  for Indice := 0 to TOBModele.detail.count-1 do
  begin
    TOBL := TOBModele.detail[Indice]; if TOBL = nil then break;
    TOBL.putValue('BSM_NUMORDRE',Indice+1);
  end;
end;

procedure TOF_BTMODELHEBDO.AfficheLagrille;
var Indice : integer;
begin
  for Indice := 0 to TOBModele.detail.count -1 do
  begin
    AfficheLigne(GS,TOBModele.detail[Indice],Indice);
  end;
end;

procedure TOF_BTMODELHEBDO.BeforeValide;
var Indice : integer;
begin
  Indice := 0;
  repeat
    if IsLigneVide (Indice+1) then
    begin
      TOBmodele.detail[Indice].free;
    end else
    begin
      TOBmodele.detail[Indice].putValue('BSM_TYPERESSOURCE',TOBmodele.GetValue('BMS_TYPERESSOURCE'));
      TOBmodele.detail[Indice].putValue('BSM_CODEMODELE',TOBmodele.GetValue('BMS_CODEMODELE'));
      TOBModele.detail[Indice].SetAllModifie (True);
      Inc(Indice);
    end;
  until Indice >= TOBModele.detail.count;
end;

procedure TOF_BTMODELHEBDO.ValideLeModele;
begin
  detruitAncienModele;
  EcritModele;
end;

procedure TOF_BTMODELHEBDO.detruitAncienModele;
var REq : string;
begin
  Req := 'DELETE FROM BTMODELSHEBDO WHERE BMS_TYPERESSOURCE="'+TOBModele.getValue('BMS_TYPERESSOURCE')+'" AND '+
         'BMS_CODEMODELE="'+TOBModele.getValue('BMS_CODEMODELE')+'"';
  ExecuteSql (Req);
  Req := 'DELETE FROM BTDETMODELSHEB WHERE BSM_TYPERESSOURCE="'+TOBModele.getValue('BMS_TYPERESSOURCE')+'" AND '+
         'BSM_CODEMODELE="'+TOBModele.getValue('BMS_CODEMODELE')+'"';
  ExecuteSql (Req);
end;

procedure TOF_BTMODELHEBDO.EcritModele;
BEGIN
  TOBModele.SetAllModifie (True);
  if not TOBModele.InsertDBByNivel (True) then V_PGI.ioError := OeUnknown;
END;

procedure TOF_BTMODELHEBDO.NextModele;
begin
  InitGrille;
  TOBModele.clearDetail;
  BMS_CODEMODELE.Text := '';
  BMS_LIBELLE.Text := '';
  SetModeSaisieGrid (false);
  EventGrille(false);
  SetMode (TmaModif);
  SetButton (True);
  BMS_CODEMODELE.SetFocus ;
end;

procedure TOF_BTMODELHEBDO.BChercheClick(Sender: TOBject);
begin

  if TFvierge(ecran).ActiveControl.Name <> 'GS' Then
  begin
  	NextControl (TFvierge(ecran),true);
  end;

  if not ControleEntete then exit;
  //
  if (ModeAction=TmaModif) then
  begin
    if not LoadModele (TOBModele ,BMS_CODEMODELE.text,BMS_TYPERESSOURCE.Value) then
    begin
      BMS_LIBELLE.text :=''; Exit;
    end else
    begin
      BMS_Libelle.text := TOBModele.getValue('BMS_LIBELLE');
    end;
  end else
  begin
    if ValidelaCreation then
    begin
      TOBModele.putValue('BMS_TYPERESSOURCE',BMS_TYPERESSOURCE.Value);
      TOBModele.putValue('BMS_CODEMODELE',BMS_CODEMODELE.text);
      TOBModele.putValue('BMS_LIBELLE',BMS_LIBELLE.text);
    end else
    begin
      Exit;
    end;
  end;
  // tutti va bene
  SetButton (false);
  ChargeLaGrille;
end;

procedure TOF_BTMODELHEBDO.SetButton (Entete : boolean);
begin
  Bannul.visible := not Entete;
  BValider.Visible := Not Entete;
  BCherche.visible := Entete;
  BFerme.visible := Entete;
end;

procedure TOF_BTMODELHEBDO.BAnnuleClick(Sender: TOBject);
begin
  if TOBmodele.detail.count > 0 then
  begin
    if PgiAsk (TraduireMemoire('Une saisie est en cours. Abandonner ?'))=MrYes then
    begin
      NextModele;
    end;
  end;
end;

function TOF_BTMODELHEBDO.ControleEntete : boolean;
begin
	result := true;
  if (ModeAction=TmaModif) then
  begin
    if BMS_CODEMODELE.text = '' then exit;
    if  (Not IsModeleExist (BMS_CODEMODELE.Text,BMS_TYPERESSOURCE.Value)) then
    begin
      PgiBox(TraduireMemoire('Ce modèle n''existe pas'));
      BMS_CODEMODELE.Text := '';
      SetFocusControl('BMS_CODEMODELE');
      result := false;
      exit;
    end;
  end else
  begin
    // Creation
    if GetControltext('BMS_TYPERESSOURCE') = '' then
    begin
      PgiBox ('Veuillez renseigner le type de ressource du modèle');
      SetFocusControl('BMS_TYPERESSOURCE');
      result := false;
      exit;
    end;
    if (IsModeleExist (BMS_CODEMODELE.Text,BMS_TYPERESSOURCE.Value)) then
    begin
      PgiBox(TraduireMemoire('Ce modèle existe déjà.'));
      BMS_CODEMODELE.Text := '';
      SetFocusControl('BMS_CODEMODELE');
      result := false;
      exit;
    end;
    if BMS_LIBELLE.Text = '' then
    begin
      PgiBox ('Veuillez renseigner la désignation du modèle');
      SetFocusControl('BMS_LIBELLE');
      result := false;
      exit;
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTMODELHEBDO ] ) ;
end.

