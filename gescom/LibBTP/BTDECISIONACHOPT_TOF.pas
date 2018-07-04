{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDECISIONACHOPT ()
Mots clefs ... : TOF;BTDECISIONACHOPT
*****************************************************************}
Unit BTDECISIONACHOPT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     fe_main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     M3Fp,
     AglInit,
     HEnt1,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     HTB97,
     Windows,
     hPanel,
     vierge,
     UTOF ;

Type
  TOF_BTDECISIONACHOPT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    ListeSaisie : string;
    TOBOptions: TOB;
    TOBdecisionnel : TOB;
    GERESTOCK,LIVCHANTIER,CHGOPT : TcheckBox;
    OPTIONS : TGRoupBox;
    FOURN : THedit;
    DEPOT : THValComboBox;
    TLIBELLEFOU : THLabel;
    BCONSULFOU,BDELETE,BFOURNDEF : TToolbarButton97;
    TLIBSTOCK,TLIBLIVR : THlabel;
    GS : THgrid;
    PANELBAS,PANELMID,PDEBUT : THpanel;
    GS_NUM,GS_CODE,GS_LIBELLE : integer;
    CellCur : string;
  	Action : TactionFiche;
    Bvalide : TToolbarButton97;
    procedure BConsultFouClick (Sender : Tobject);
    procedure GetControls;
    procedure SetControls;
    procedure SetEvents;
    procedure FournisseurExit (Sender : TObject);
    procedure FournElipsisClick(Sender: TObject);
    procedure DefiniGrilleSaisie;
    function AddLigneConsult (WithNumerotation : boolean = false): TOB;
    procedure AddChampsLigCOnsult(TOBL: TOB);
    procedure NumeroteConsult;
    procedure AfficheLaGrille;
    procedure AfficheLaLigne(TOBL: TOB; Ligne: integer);
    procedure InitTOB;
    procedure GSEnter(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);

    procedure GSElipsisClick (Sender : TObject);
    function  ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
    procedure SetEventGrid(Activation: boolean);
    function LigneVide(Arow : integer): boolean;
    function ControleLigne(Ou: integer): boolean;
    procedure PositionneDansGrid(Arow, Acol: integer);
    procedure ChangeFournisseur(Arow: integer; NewValue: String;
      var Cancel: boolean);
    function ExistFournisseur(Valeur: string;
      var libelle: string): boolean;
    procedure FournSelElipsisClick(Sender: TObject);
    procedure BvalideClick (Sender : Tobject);
    procedure ShowButtons(status: boolean);
    procedure BdeleteClick(Sender: TObject);
    procedure BFournDefClick(Sender: TObject);
    procedure AffecteLeFournisseur(TOBD: TOB);
    procedure AjouteLesCatalogues(TheTOBCons: TOB);
    procedure ChgOptClick(Sender: TObject);
    procedure GERESTOCKClick (Sender : TObject);
    procedure LIVCHANTIERClick (Sender : TObject);

  end ;

Implementation
uses DECISIONACHAT_TOF;


procedure TOF_BTDECISIONACHOPT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDECISIONACHOPT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDECISIONACHOPT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTDECISIONACHOPT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDECISIONACHOPT.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBOptions := LaTOB; // recup des infos
  TOBDecisionnel := TOB(TOBOptions.data);
  GetControls;
  ecran.height := 154;
  SetControls;
  SetEvents;
end ;

procedure TOF_BTDECISIONACHOPT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDECISIONACHOPT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDECISIONACHOPT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTDECISIONACHOPT.GetControls;
begin
  GERESTOCK := TCheckBox(GetControl('GERESTOCKS'));
  LIVCHANTIER := TcheckBox(GetControl('LIVCHANTIERS'));
  //
  TLIBSTOCK := THlabel (GetCOntrol('TLIBSTOCK'));
  TLIBLIVR := THlabel (GetCOntrol('TLIBLIVR'));
  //
  CHGOPT := TcheckBox(GetControl('CHGOPT'));
  OPTIONS := TGRoupBox(GetControl('OPTIONS'));

  //
  FOURN := THedit(GetControl('FOURN'));
  DEPOT := THValComboBox(GetControl('DEPOT'));
  TLIBELLEFOU := THLabel(GetCOntrol('TLIBELLEFOU'));
  BCONSULFOU := TToolbarButton97(Getcontrol('BCONSULTFOU'));
  BDELETE := TToolbarButton97(Getcontrol('BDELETE'));
  BFOURNDEF := TToolbarButton97(Getcontrol('BFOURNDEF'));
  GS := ThGrid(GetControl('GS'));
  PANELBAS := ThPanel(GetControl('PANELBAS'));
  PANELMID := ThPanel(GetControl('PANELMID'));
  PDEBUT := ThPanel(GetControl('PDEBUT'));
  Bvalide := TToolbarButton97 (GetControl('BVALIDER'));
end;

procedure TOF_BTDECISIONACHOPT.SetControls;
begin
  BCONSULFOU.Visible := (TOBOptions.GetValue('CONSULTFOU')='X');
  ShowButtons (false);
  TLIBELLEFOU.Caption := '';
  PANELMID.visible := false;
  PANELBAS.visible := false;
  GERESTOCK.state := CbGrayed;
  LIVCHANTIER.state := CbGrayed;
end;

procedure TOF_BTDECISIONACHOPT.SetEvents;
begin
  FOURN.OnElipsisClick := FournSelElipsisClick;
  Fourn.OnExit := FournisseurExit;
  BCONSULFOU.onClick := BConsultFouClick;
  BDELETE.onClick := BdeleteClick;
  BFOURNDEF.OnCLick := BFournDefClick;
  CHGOPT.OnClick := ChgOptClick;
  GERESTOCK.onClick := GERESTOCKClick;
  LIVCHANTIER.onClick := LIVCHANTIERClick;
  GERESTOCKClick (self);
  LIVCHANTIERClick (self);
  //
  Bvalide.onClick := BvalideClick;
  //
end;

procedure TOF_BTDECISIONACHOPT.FournSelElipsisClick (Sender : TObject);
var MaCle,Libelle : string;
begin
  MaCle := '';
  MaCle := AGLLanceFiche('GC', 'GCFOURNISSEUR_MUL', 'T_NATUREAUXI=FOU', '', 'SELECTION');
  if (MaCle<>'') and (MaCle <> FOurn.Text) then
  begin
    if ExistFournisseur (MaCle,Libelle) then
    begin
      Fourn.text := MaCle;
      TLIBELLEFOU.Caption := Libelle;
      CellCur := GS.Cells [GS.Col,GS.row];
    end;
  end;
end;

procedure TOF_BTDECISIONACHOPT.FournElipsisClick (Sender : TObject);
var TOBRecupFou,LAlIG : TOB;
    Indice : integer;
    SQL : String;
    QQ : Tquery;
begin
  TOBRecupFou := TOB.Create ('LES FOURNISSEURS',nil,-1);
  TRY
    TheTOB := TOBrecupFou;
    AGLLanceFiche('BTP', 'BTFOURNMULTI_MUL', 'T_NATUREAUXI=FOU', '', 'SELECTION');
    if TheTOB = nil then exit;
    if TheTOB.detail.count = 0 then exit;
    for Indice := 0 to TheTOB.detail.count -1 do
    begin
      if TOBOptions.findFirst(['FOURNISSEUR'],[TheTOB.detail[Indice].getValue('T_TIERS')],True) = nil then
      begin
        // ajoute les lignes founisseurs
        if Indice = 0 Then
        begin
          LAlIG := TOBOptions.detail[GS.row-1];
        end else
        begin
          LAlIG := AddLigneConsult(TRue);
        end;
        Sql := 'SELECT T_TIERS,T_LIBELLE FROM TIERS WHERE T_TIERS="'+TheTOB.detail[Indice].getValue('T_TIERS')+'" AND T_NATUREAUXI="FOU"';
        QQ := OPenSql (Sql,True);
        LAlIG.putValue('FOURNISSEUR',QQ.findField('T_TIERS').AsString);
        LAlIG.putValue('LIBELLE',QQ.findField('T_LIBELLE').AsString);
        ferme (QQ);
      end;
    end;
  FINALLY
    TheTOB := nil;
    TOBRecupFou.free;
  END;
  AfficheLaGrille ;
  CellCur:=GS.Cells[GS.col,GS.row] ;
end;


procedure TOF_BTDECISIONACHOPT.FournisseurExit(Sender: TObject);
var Libelle : string;
begin
  if Fourn.text = '' then BEGIN TLIBELLEFOU.caption := ''; exit; END;
  if ExistFournisseur (Fourn.text,Libelle) then
  begin
    TLIBELLEFOU.caption := Libelle;
  end else
  begin
    PGIError ('Fournisseur Inconnu...');
    TLIBELLEFOU.caption := '';
    Fourn.text := '';
    Fourn.SetFocus ;
  end;
end;


procedure TOF_BTDECISIONACHOPT.AddChampsLigCOnsult (TOBL : TOB);
begin
  TOBL.AddChampSupValeur ('NUM',0);
  TOBL.AddChampSupValeur ('FOURNISSEUR','');
  TOBL.AddChampSupValeur ('LIBELLE','');
end;

function TOF_BTDECISIONACHOPT.AddLigneConsult (WithNumerotation : boolean = false): TOB;
var TOBL : TOB;
begin
  TOBL := TOB.Create ('LIGCONSUL',TOBOptions,-1);
  AddChampsLigCOnsult (TOBL);
  if WithNumerotation then TOBL.putValue('NUM',TOBOptions.detail.count);
  result := TOBL;
end;

procedure TOF_BTDECISIONACHOPT.NumeroteConsult;
var Indice : integer;
    TOBL : TOB;
begin
  for Indice := 0 to TOBOptions.detail.count -1 do
  begin
    TOBL := TOBOptions.detail[indice];
    TOBL.PutValue('NUM',Indice+1);
  end;
end;

procedure TOF_BTDECISIONACHOPT.BConsultFouClick(Sender: Tobject);
var cancel : boolean;
    Arow,Acol : integer;
begin
  if BCONSULFOU.Down then
  begin
    ecran.height := 290;
    PDEBUT.enabled := false;
    PANELMID.visible := true;
    PANELBAS.visible := true;
    ShowButtons(true);
    ListeSaisie := 'NUM;FOURNISSEUR;LIBELLE';
    DefiniGrilleSaisie;
    InitTOB;
    AfficheLaGrille;
    TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
    TLIBELLEFOU.caption := '';
    Fourn.text := '';
    GS.SetFocus;
    GSEnter (self);
    SetEventGrid (True);
  end else
  begin
    if PGIAsk ('Désirez-vous réellement annuler la consultation fournisseur ?')=Mryes then
    begin
      ShowButtons(false);
      SetEventGrid (false);
      TOBOptions.clearDetail;
      GS.VidePile (false);
      PANELMID.visible := false;
      PANELBAS.visible := false;
      PDEBUT.enabled := true;
      ecran.height := 154;
    end else
    begin
      BCONSULFOU.Down := true;
    end;
  end;
end;

procedure TOF_BTDECISIONACHOPT.DefiniGrilleSaisie;
var Numero : integer;
		TheListe,Unchamp  : string;
begin
  TheListe := ListeSaisie;
  numero := 0;
  repeat
  	UnChamp := READTOKENST (TheListe);
    if UnChamp <> '' then
    begin
      if UnChamp = 'NUM' then GS_NUM := numero else
      if UnChamp = 'FOURNISSEUR' then GS_CODE := numero else
      if UnChamp = 'LIBELLE' then GS_LIBELLE := numero;
      inc(numero);
    end;
  until UnChamp='';

  GS.ColCount := Numero;

  GS.cells[GS_NUM,0] := '';
  GS.ColWidths [GS_NUM] := 18;

  // Affichage du Code fournisseur
  GS.cells[GS_CODE,0] := 'Fournisseur';
  GS.ColWidths [GS_CODE] := 120;

  // Affichage du Libelle fournisseur
  GS.cells[GS_LIBELLE,0] := 'Désignation';
  GS.ColWidths [GS_LIBELLE] := 300;
	GS.ColAligns[GS_LIBELLE]:=taLeftJustify;
end;

procedure TOF_BTDECISIONACHOPT.AfficheLaLigne (TOBL : TOB; Ligne: integer);
begin
  TOBL.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);
end;

procedure TOF_BTDECISIONACHOPT.AfficheLaGrille;
var indice: integer;
    TOBL : TOB;
begin
  SetEventGrid (false);
	Gs.SynEnabled := false;
  GS.BeginUpdate;
  GS.rowCount := TOBOptions.detail.count + 1;
  for Indice := 0 to TOBOptions.detail.count -1 do
  begin
    TOBL := TOBOptions.detail[Indice];
    AfficheLaLigne (TOBL,Indice+1);
  end;
  gs.EndUpdate;
  gs.SynEnabled := true;
  SetEventGrid (True);
end;

procedure TOF_BTDECISIONACHOPT.InitTOB;
begin
  if TOBOptions.detail.count > 0 then TOBOptions.ClearDetail;
  AddLigneConsult;
  NumeroteConsult;
end;

procedure TOF_BTDECISIONACHOPT.GSEnter(Sender: TObject);
var Acol,Arow : integer;
    cancel : boolean;
begin
  cancel := false;
  GS.row := 1;
  GS.col := 1;
  Arow := 1;
  Acol := 1;
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  cancel := false;
  GSRowEnter (GS,Arow,cancel,false);
  GSCellEnter (GS,Acol,Arow,cancel);
  GS.row := Arow;
  GS.col := Acol;
  GS.ShowEditor;
  GS.ElipsisButton := (Acol = GS_CODE);
  CellCur := GS.Cells [Acol,Arow];
end;

procedure TOF_BTDECISIONACHOPT.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  if Action=taConsult then Exit ;
  ZoneSuivanteOuOk(ACol,ARow,Cancel) ;
  GS.ElipsisButton := (Acol=GS_CODE);
  if not cancel then
  begin
    CellCur:=GS.Cells[ACol,ARow] ;
  end;
end;

procedure TOF_BTDECISIONACHOPT.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
	if Action = TaConsult then exit;
  if CellCur = GS.cells[Acol,Arow] then exit;
  TOBL := TOBOptions.detail[Arow-1];
  if (Acol <> GS_CODE) then
  begin
    GS.Cells[Acol,Arow] := CellCur;
    exit;
  end;
  //
  if Acol = GS_CODE then
  begin
  	ChangeFournisseur (Arow,GS.Cells[Acol,Arow],Cancel);
  end;

end;

procedure TOF_BTDECISIONACHOPT.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
  if (Ou > 1) and (LigneVide (Ou-1)) then
  BEGIN
    cancel := true;
    Exit;
  END;
  if (Ou >= GS.rowCount -1) then GS.rowCount := GS.rowCount +1;
  if Ou > TOBOptions.Detail.count then
  begin
    TOBL := AddLigneConsult (True);
    AfficheLaLigne  (TOBL,Ou);
  end;
end;

procedure TOF_BTDECISIONACHOPT.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var Acol : integer;
begin
  if Action = taConsult then Exit;

  // si vide pas besoin de controler le contenu
  if LigneVide (Ou) then exit;

  if not ControleLigne (Ou) then
  BEGIN
  	cancel := true;
    Acol := GS_CODE;
    PositionneDansGrid (Ou,Acol);
  END;
end;

function TOF_BTDECISIONACHOPT.ZoneAccessible(ACol, ARow: Integer): boolean;
Var TOBL : TOB ;
BEGIN
  Result:=True ;
  if (Acol <> GS_CODE) then
  begin
  	result := false;
    exit;
  end;
end;

procedure TOF_BTDECISIONACHOPT.ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
Var Sens,ii,Lim : integer ;
    OldEna,ChgLig,ChgSens  : boolean ;
BEGIN
  OldEna:=GS.SynEnabled ; GS.SynEnabled:=False ;
  Sens:=1 ; ChgLig:=(GS.Row<>ARow) ; ChgSens:=False ;
  ACol:=GS.Col ; ARow:=GS.Row ; ii:=0 ;
  While Not ZoneAccessible(ACol,ARow)  do
  BEGIN
    Cancel:=True ; inc(ii) ; if ii>500 then Break ;
    if Sens=1 then
    BEGIN
      Lim:=GS.rowCount-1;
      if ((ACol=GS.ColCount-1) and (ARow>=Lim)) then
      BEGIN
//        if ChgSens then Break else BEGIN Sens:=-1 ; Continue ; ChgSens:=True ; END ;
        Break  ;
      END ;
      if ChgLig then BEGIN ACol:=GS_CODE-1 ; ChgLig:=False ; END ;
      if ACol<GS.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=GS_CODE ;END ;
    END else
    BEGIN
      if ((ACol=GS_CODE) and (ARow=1)) then
      BEGIN
        if ChgSens then Break else BEGIN Sens:=1 ; Continue ; END ;
      END ;
      if ChgLig then BEGIN ACol:=GS_CODE+2; ChgLig:=False ; END ;
      if ACol>GS_CODE then Dec(ACol) else BEGIN Dec(ARow) ; ACol:=GS_CODE+2 ; END ;
    END ;
  END ;
  GS.SynEnabled:=OldEna ;
end;

procedure TOF_BTDECISIONACHOPT.GSElipsisClick(Sender: TObject);
begin
  FournElipsisClick (self);
end;

procedure TOF_BTDECISIONACHOPT.SetEventGrid (Activation : boolean);
begin
  if Activation then
  begin
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.OnElipsisClick := GSElipsisClick;
  end else
  begin
    GS.OnCellEnter := nil;
    GS.OnCellExit := nil;
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnElipsisClick := nil;
  end;
end;


function TOF_BTDECISIONACHOPT.LigneVide (Arow : integer) : boolean;
var TOBL : TOB;
begin
  if ARow > TOBOptions.detail.count then BEGIN result := true; exit; END;
  TOBL := TOBOptions.detail[Arow-1];
  result := (TOBL.GetValue('FOURNISSEUR')='');
end;

function TOF_BTDECISIONACHOPT.ControleLigne (Ou : integer) : boolean;
var TOBL : TOB;
    Valeur,Libelle : string;
begin
  result := false;
  if Ou > TOBOPtions.detail.count then exit;
  TOBL := TOBOPtions.detail[Ou-1];
  Valeur := TOBL.GetValue('FOURNISSEUR');
  result := ExistFournisseur (Valeur,Libelle);
end;

procedure TOF_BTDECISIONACHOPT.PositionneDansGrid (Arow,Acol : integer);
var cancel : boolean;
begin
  GS.CacheEdit;
  SetEventGrid (false);
  GS.row := Arow;
  GS.col := Acol;
  SetEventGrid (true);
  GSRowEnter (self,Arow,Cancel,false);
  GSCellEnter (self,Acol,Arow,Cancel);
  GS.row := Arow;
  GS.col := Acol;
  GS.ShowEditor;
end;

function TOF_BTDECISIONACHOPT.ExistFournisseur (Valeur : string; var libelle : string) : boolean;
var QQ : TQuery;
begin
  Libelle := '';
  QQ := OPenSql ('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+Valeur+'" AND T_NATUREAUXI="FOU"',True);
  result := not QQ.eof;
  if not QQ.eof then Libelle := QQ.FindField('T_LIBELLE').asString;
  ferme (QQ);
end;

procedure TOF_BTDECISIONACHOPT.ChangeFournisseur(Arow: integer;NewValue: String; var Cancel : boolean);
var TOBL : TOB;
		NBligne , Ligne : integer;
    Libelle : string;
begin
	cancel := false;
	if (not ExistFournisseur(newvalue,Libelle)) and (NewValue <> '') then
  begin
  	PGIBox(TraduireMemoire('Ce fournisseur n''existe pas..'),ecran.caption);
    cancel := true;
    exit;
  end;

  if TOBOptions.findFirst(['FOURNISSEUR'],[newValue],True) <> nil then
  begin
  	PGIBox(TraduireMemoire('Fournisseur déjà indiqué..'),ecran.caption);
    cancel := true;
    exit;
  end;

  TOBL := TOBOptions.detail[Arow-1];
  Ligne := TOBL.GetValue('NUM');
  TOBL.PutValue('FOURNISSEUR',NewValue);
  TOBL.PutValue('LIBELLE',Libelle);
  AfficheLaLigne (TOBL,Arow);
  CellCur:=GS.Cells[GS.col,GS.row] ;
end;

procedure TOF_BTDECISIONACHOPT.BvalideClick(Sender: Tobject);
var libelle : string;
    Indice : integer;
begin
  if BCONSULFOU.down then
  begin
    // application des consultations fournisseurs
    Indice := 0;
    repeat
      if LigneVide (Indice+1) then TOBOptions.detail[Indice].free else Inc(Indice);
    until Indice >= TOBOptions.detail.count;
    TheTOB := TOBOptions;
  end else
  begin
    // Modification des parametres
    if (Fourn.Text <> '') and (not ExistFournisseur (Fourn.text,Libelle)) then
    begin
      PGIError ('Fournisseur Inconnu...');
      TLIBELLEFOU.caption := '';
      Fourn.text := '';
      Fourn.SetFocus ;
      TFVierge(Ecran).ModalResult := 0;
      exit;
    end else
    begin
      TOBOptions.PutValue('FOURN',Fourn.text);
      TOBOptions.PutValue('DEPOT',Depot.value);
      if CHGOPT.Checked then
      begin
        TOBOptions.PutValue('CHGOPT','X');
        if LIVCHantier.State = cbChecked then
        begin
          TOBOptions.PutValue('LIVCHANTIER','X')
        end else if LIVCHantier.State = cbUnChecked then
        begin
          TOBOptions.PutValue('LIVCHANTIER','-');
        end else
        begin
          TOBOptions.PutValue('LIVCHANTIER','');
        end;
        if GERESTOCK.State=CbChecked then
        begin
          TOBOptions.PutValue('GERESTOCK','X')
        end else if GERESTOCK.State=CbUnChecked then
        begin
          TOBOptions.PutValue('GERESTOCK','-') ;
        end else
        begin
          TOBOptions.PutValue('GERESTOCK','') ;
        end;
      end;
      TheTOB := TOBOptions;
    end;
  end;
end;


procedure TOF_BTDECISIONACHOPT.BdeleteClick (Sender : TObject);
begin
  if PgiAsk('Etes-vous sur de vouloir supprimer tous les fournisseurs de la consultation ?')=Mryes then
  begin
    InitTOB;
    GS.VidePile (false);
    AfficheLaGrille;
    TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
    TLIBELLEFOU.caption := '';
    Fourn.text := '';
    GS.SetFocus;
    GSEnter (self);
    SetEventGrid (True);
  end;
end;

procedure TOF_BTDECISIONACHOPT.AffecteLeFournisseur (TOBD : TOB);
var leFourn,Libelle : string;
    TOBL : TOB;
    Acol,Arow : integer;
    Cancel : boolean;
begin
  leFourn := TOBD.GetValue('BAD_FOURNISSEUR');
  if Lefourn <> '' then
  begin
    if TOBOptions.findFirst(['FOURNISSEUR'],[lefourn],True) <> nil then
    begin
      TOBL := AddLigneConsult(True);
      ExistFournisseur (leFourn,Libelle);
      TOBL.PutValue('FOURNISSEUR',LeFourn);
      TOBL.PutValue('LIBELLE',Libelle);
    end;
  end;
  //
  AjouteLesCatalogues (TOBD);
end;

procedure TOF_BTDECISIONACHOPT.BFournDefClick (Sender : TObject);
var Indice : integer;
    TOBD : TOB;
    Acol,Arow : integer;
begin
  Indice := 0;
  Repeat
    if TOBOptions.detail[Indice].getValue('FOURNISSEUR')='' then TOBOptions.detail[Indice].free;
  until Indice >= TOBOptions.detail.count;
  for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD:= GetSelectionne (TOBDecisionnel.detail[Indice]);
    if TOBD <> nil then
    begin
      AffecteLeFournisseur (TOBD);
    end;
  end;
  //
  // affichage de la grille
  if TOBOptions.detail.count = 0 then AddLigneConsult (True);
  AfficheLaGrille;
  Arow := GS.row;  if Arow < 1 then Arow := 1;
  Acol := GS_CODE;
  PositionneDansGrid (Arow,Acol);
end;

procedure TOF_BTDECISIONACHOPT.ShowButtons(status : boolean);
begin
  BDELETE.visible := status;
  BFOURNDEF.visible := status;
end;


procedure TOF_BTDECISIONACHOPT.AjouteLesCatalogues (TheTOBCons : TOB);
var TOBcata,TOBL : TOB;
    Req,lefourn,Libelle : String;
    Indice : integer;
begin
  TOBcata := TOB.Create ('LES CATALOG',nil,-1);
  Req := 'SELECT *,T_LIBELLE FROM CATALOGU LEFT JOIN TIERS ON T_TIERS=GCA_TIERS AND T_NATUREAUXI="FOU" WHERE GCA_ARTICLE="'+
          TheTOBCons.getValue('BAD_ARTICLE')+'"';
  TOBCata.LoadDetailDBFromSQL ('CATALOGU',req,false);
  for Indice := 0 to TOBCata.detail.count -1  do
  begin
    leFourn := TOBCata.detail[Indice].GetValue('GCA_TIERS');
    if TOBOptions.findFirst(['FOURNISSEUR'],[lefourn],True) = nil then
    begin
      TOBL := AddLigneConsult(True);
      TOBL.putValue('FOURNISSEUR',leFourn);
      TOBL.putValue('LIBELLE',TOBCata.detail[Indice].GetValue('T_LIBELLE'));
    end;
  end;
  TOBcata.free;
end;

procedure TOF_BTDECISIONACHOPT.ChgOptClick (Sender : TObject);
begin
  OPTIONS.Enabled := CHGOPT.Checked; 
end;

procedure TOF_BTDECISIONACHOPT.GERESTOCKClick(Sender: TObject);
begin
  if GERESTOCK.state = CbGrayed then
  begin
    TLIBSTock.Caption := 'Pas de Changement'
  end else
  begin
    TLIBSTock.Caption := ''
  end;
end;

procedure TOF_BTDECISIONACHOPT.LIVCHANTIERClick(Sender: TObject);
begin
  if LIVCHANTIER.state = CbGrayed then
  begin
    TLIBLIVR.Caption := 'Pas de Changement'
  end else
  begin
    TLIBLIVR.Caption := ''
  end;
end;

Initialization
  registerclasses ( [ TOF_BTDECISIONACHOPT ] ) ;
end.

