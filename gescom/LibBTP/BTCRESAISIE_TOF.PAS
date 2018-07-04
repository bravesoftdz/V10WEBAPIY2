{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/01/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCRESAISIE ()
Mots clefs ... : TOF;BTCRESAISIE
*****************************************************************}
Unit BTCRESAISIE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HPanel,
     HMsgBox,
     HTB97,
     UentCommun,
     Windows,
     Messages,
     UTOF ;


Type

  TChps = class
    Nomchps : string;
    TypeChps : string;
    LibelleINI : string;
    ModifINI : boolean;
    TailleINI : integer;
    NbDecIni : integer;
    AlignINI : TCaligment;
    //
    Modifiable : boolean;
    Designation : string;
    TailleReele : integer;
    NbDec : integer;
    Align : TCaligment;
    SepMilliers : Boolean;
    LibComplet : boolean;
    Obligatoire : Boolean;
    visible : boolean;
    BlancSiNul : boolean;
    WithCumul : boolean;
  end;

  TlistChps = class (Tlist)
  private
    function Add(AObject: TChps): Integer;
    function GetItems(Indice: integer): TChps;
    procedure SetItems(Indice: integer; const Value: TChps);
  public
    destructor destroy; override;
    property Items [Indice : integer] : TChps read GetItems write SetItems;
    function find (NomChps : string): TChps;
  end;


  TOF_BTCRESAISIE = Class (TOF)
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
    TOBLISTE : TOB;
    fIndexItem : integer;
    fNewListe : boolean;
    fDataModif : boolean;
    fDataChanged : boolean;
    fListeSaisie : string;
    LCHAMPS : TListBox;
    LLIBELLES : TListBox;
    LCHPSELECT : TlistBox;
    LLIBSELECT : TlistBox;
    BIN,BOUT,BUP,BDOWN : TToolbarButton97;
    RBCHPS,RBLIBELLE : TRadioButton;
    //
    ZCODELISTE : Thedit;
    ZNOMLISTE : Thedit;
    ZTITLE : Thedit;
    ZLARG : THSpinEdit;
    ZNBDEC : THSpinEdit;
    ZSEPMIL : ThCheckBox;
    ZVIDESINUL : THCheckbox;
    BALLEFT,BALCENTER,BALRIGHT : TToolbarButton97;
    ZVISIBLE : ThCheckBox;
    ZCUMUL : ThCheckBox;
    ZMODIFOK : ThCheckBox;
    ZOBLIG : THCheckbox;
    //
    LaListeChps : TlistChps;
    TS: TStringList;
    procedure RemplitLaliste;
    procedure SetListChamps;
    procedure GetControls;
    procedure SetEvents(Status: boolean);
    procedure BInClick (Sender : Tobject);
    procedure BOutClick (Sender : Tobject);
    procedure ChangeModeSel(Sender: Tobject);
    procedure SetShow;
    procedure DecodeArguments (Arg : string);
    procedure SelectionneThisItem(TheItem: TChps);
    procedure SetInfoListeRecup;
    procedure GetInfoListe;
    procedure ChpsSelect (Sender : Tobject);
    procedure LibSelect (Sender : Tobject);
    procedure CHPSELClick (Sender : Tobject);
    procedure LIBSELClick (Sender : Tobject);
    procedure AfficheInfos (SelItem : TChps);
    procedure BUpClick (Sender : Tobject);
    procedure BDownClick (Sender : Tobject);
    procedure EnregInfo;
    procedure DataChange (Sender : Tobject);
    procedure EnregistreListe;
    procedure InitEcran;
    function ConstitueLigne(TtypeDataExt: TListeTypeDat): string;

  end ;

Implementation

uses TntStdCtrls,EntGC;

function CompareNames(Item1, Item2: Pointer): Integer;
begin
  Result := CompareText(TChps(Item1).Nomchps, TChps(Item2).Nomchps);
end;

procedure TOF_BTCRESAISIE.OnNew ;
var Indice : integer;
begin
  if fDataChanged then
  begin
    if PgiAsk ('ATTENTION : Vous allez perdre les modifications.#13#10 Confirmez-vous l''abandon ?') <> mrYes then
    begin
      ecran.ModalResult := 0;
      exit;
    end;
  end;
  Inherited ;
  (*
  for Indice := LaListeChps.count -1 downto 0 do
  begin
    if TChps(LaListeChps.Items [Indice])<> nil then TChps(LaListeChps.Items [Indice]).free;
  end;
  LaListeChps.Clear;
  LCHAMPS.clear;
  LLIBELLES.clear;
  TOBLISTE.InitValeurs;
  RemplitLaliste;
  LCHPSELECT.Clear;
  LLIBSELECT.Clear;
  SetListChamps;
  *)
  fNewListe := true;
  fDataModif := false;
  fDataChanged := true;
  TToolbarButton97 (GetControl('Binsert')).enabled := false;
  TToolbarButton97 (GetControl('Bdelete')).enabled := false;
  InitEcran;
end ;

procedure TOF_BTCRESAISIE.OnDelete ;
var CodeListe : string;
begin
  Inherited ;
  CodeLIste := ThEdit(GetControl('ZCODELISTE')).Text;
  if Copy(CodeListe,1,1)<> 'Z' then
  begin
    PGiInfo ('Vous ne pouvez pas supprimer les listes apportées par L.S.E');
    ecran.ModalResult := 0;
    exit;
  end;
  if PgiAsk('ATTENTION : Vous allez supprimer cette liste.#13#10 Confirmez-vous ?')=MrYes then
  begin
    ExeCuteSql('DELETE FROM LISTE WHERE LI_LISTE="'+CodeListe+'"');
    PostMessage(TFmul(Ecran).Handle, WM_CLOSE, 0, 0);
  end;
end ;

procedure TOF_BTCRESAISIE.OnUpdate ;
var CodeListe : string;
begin
  Inherited ;
  if fIndexItem >= 0 then
  begin
    EnregInfo;
  end;

  if not fDataChanged then exit;

  if LCHPSELECT.Items.Count = 0 then
  begin
    PGiInfo ('Vous devez indiquer des champs dans la liste de saisie');
    ecran.ModalResult := 0;
    exit;
  end;
  if fNewListe then
  begin
    CodeLIste := ThEdit(GetControl('ZCODELISTE')).Text;
    if CodeLIste = '' then
    begin
      PGiInfo ('Merci de renseigner un code de liste');
      ecran.ModalResult := 0;
      ThEdit(GetControl('ZCODELISTE')).SetFocus;
      exit;
    end;
    if CodeLIste = 'Z' then
    begin
      PGiInfo ('Merci de renseigner un code de liste');
      ecran.ModalResult := 0;
      ThEdit(GetControl('ZCODELISTE')).SetFocus;
      exit;
    end;
    if copy(CodeLIste,1,1)<> 'Z' then
    begin
      PGiInfo ('les listes spécifiques doivent commencer par "Z"');
      ecran.ModalResult := 0;
      ThEdit(GetControl('ZCODELISTE')).SetFocus;
      exit;
    end;
    if ThEdit(GetControl('ZNOMLISTE')).Text = '' then
    begin
      PGiInfo ('Merci de renseigner un descrptif de la liste');
      ecran.ModalResult := 0;
      ThEdit(GetControl('ZNOMLISTE')).SetFocus;
      exit;
    end;
    if ExisteSQL('SELECT 1 FROM LISTE WHERE LI_LISTE="'+CodeListe+'"') then
    begin
      PGiInfo ('Cette liste existe déjà.');
      ecran.ModalResult := 0;
      ThEdit(GetControl('ZNOMLISTE')).SetFocus;
      exit;
  end;
  end;
  if fDataChanged then EnregistreListe;
end ;

procedure TOF_BTCRESAISIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCRESAISIE.GetControls;
begin

  LCHAMPS := TListBox(getControl('LCHAMPS'));
  LLIBELLES := TListBox(getControl('LLIBELLES'));
  LCHPSELECT := TListBox(getControl('LCHPSELECT'));
  LLIBSELECT := TListBox(getControl('LLIBSELECT'));
  BIN := TToolbarButton97 (GetControl('BIN'));
  BOUT := TToolbarButton97 (GetControl('BOUT'));
  BUP := TToolbarButton97 (GetControl('BUP'));
  BDOWN := TToolbarButton97 (GetControl('BDOWN'));
  RBCHPS := TRadioButton (GetControl('RBCHPS'));
  RBLIBELLE := TRadioButton (GetControl('RBLIBELLE'));
  //
  ZCODELISTE := Thedit(GetControl('ZCODELISTE'));
  ZNOMLISTE := Thedit(GetControl('ZNOMLISTE'));
  ZTITLE := Thedit(GetControl('ZTITLE'));
  ZLARG := THSpinEdit(GetControl('ZLARG'));
  ZNBDEC := THSpinEdit(GetControl('ZNBDEC'));
  ZSEPMIL := ThCheckBox(GetControl('ZSEPMIL'));
  ZVIDESINUL := ThCheckBox(GetControl('ZVIDESINUL'));
  BALLEFT := TToolbarButton97 (GetControl('BALLEFT'));
  BALCENTER := TToolbarButton97 (GetControl('BALCENTER'));
  BALRIGHT := TToolbarButton97 (GetControl('BALRIGHT'));
  ZVISIBLE := ThCheckBox(GetControl('ZVISIBLE'));
  ZCUMUL := ThCheckBox(GetControl('ZCUMUL'));
  ZMODIFOK := ThCheckBox(GetControl('ZMODIFOK'));
  ZOBLIG := ThCheckBox(GetControl('ZOBLIG'));

end;

procedure TOF_BTCRESAISIE.ChangeModeSel (Sender : Tobject);
var TheItem : TChps;
  IndBis : integer;
begin
  if RBCHPS.Checked then
  begin
    if LLIBELLES.ItemIndex >= 0 then
    begin
      TheItem := TChps(LLIBELLES.Items.Objects[LLIBELLES.ItemIndex]);
      IndBis := LCHAMPS.Items.IndexOf(TheItem.Nomchps);
      LCHAMPS.ItemIndex := IndBis;
    end;
  end else
  begin
    if LCHAMPS.ItemIndex >= 0 then
    begin
      TheItem := TChps(LCHAMPS.Items.Objects[LCHAMPS.ItemIndex]);
      IndBis := LLIBELLES.Items.IndexOf(TheItem.LibelleINI);
      LLIBELLES.ItemIndex := IndBis;
    end;
  end;
  SetShow;
end;

procedure TOF_BTCRESAISIE.SetEvents(Status : boolean);
begin

  if Status then
  begin
    BIN.OnClick := BINClick;
    BOUT.Onclick := BOUTClick;
    BUP.OnClick := BUpClick;
    BDOWN.OnClick := BDownClick;
    RBCHPS.OnClick := ChangeModeSel;
    RBLIBELLE.OnClick := ChangeModeSel;
    LCHAMPS.OnDblClick := BINClick;
    LLIBELLES.OnDblClick := BInClick;
    LCHPSELECT.OnDblClick := BOUTClick;
    LLIBSELECT.OnDblClick := BOUTClick;
    LCHAMPS.OnClick := ChpsSelect;
    LLIBELLES.OnClick := LibSelect;
    LCHPSELECT.OnClick := CHPSELClick;
    LLIBSELECT.OnClick := LIBSELClick;
    //
    ThEdit(GetControl('ZTITLE')).OnChange := DataChange;
    ThSpinEdit(GetControl('ZLARG')).OnChange := DataChange;
    ThSpinEdit(GetControl('ZNBDEC')).OnChange := DataChange;
    ThCheckBox(GetControl('ZSEPMIL')).OnClick := DataChange;
    TToolbarButton97(GetControl('BALLEFT')).OnClick := DataChange;
    TToolbarButton97(GetControl('BALCENTER')).OnClick := DataChange;
    TToolbarButton97(GetControl('BALRIGHT')).OnClick := DataChange;
    ThCheckBox(GetControl('ZVISIBLE')).OnClick := DataChange;
    ThCheckBox(GetControl('ZCUMUL')).OnClick := DataChange;
    ThCheckBox(GetControl('ZMODIFOK')).OnClick := DataChange;
    ThCheckBox(GetControl('ZOBLIG')).OnClick := DataChange;
    ThCheckBox(GetControl('ZVIDESINUL')).OnClick := DataChange;
    //
  end else
  begin
    BIN.OnClick := nil;
    BOUT.Onclick := nil;
    RBCHPS.OnClick := nil;
    RBLIBELLE.OnClick := nil;
    LCHAMPS.OnDblClick := nil;
    LLIBELLES.OnDblClick := nil;
    LCHPSELECT.OnDblClick := nil;
    LLIBSELECT.OnDblClick := nil;
    LCHAMPS.OnClick := nil;
    LLIBELLES.OnClick := nil;
    LCHPSELECT.OnClick := nil;
    LLIBSELECT.OnClick := nil;
    //
    ThEdit(GetControl('ZTITLE')).OnChange := nil;
    ThSpinEdit(GetControl('ZLARG')).OnChange := nil;
    ThSpinEdit(GetControl('ZNBDEC')).OnChange := nil;
    ThCheckBox(GetControl('ZSEPMIL')).OnClick := nil;
    TToolbarButton97(GetControl('BALLEFT')).OnClick := nil;
    TToolbarButton97(GetControl('BALCENTER')).OnClick := nil;
    TToolbarButton97(GetControl('BALRIGHT')).OnClick := nil;
    ThCheckBox(GetControl('ZVISIBLE')).OnClick := nil;
    ThCheckBox(GetControl('ZCUMUL')).OnClick := nil;
    ThCheckBox(GetControl('ZMODIFOK')).OnClick := nil;
    ThCheckBox(GetControl('ZOBLIG')).OnClick := nil;
    //
  end;

end;

procedure TOF_BTCRESAISIE.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBLISTE := TOB.Create ('LISTE',nil,-1);
  fIndexItem := -1;
  fDataModif := false;
  TS:= TStringList.Create;
  DecodeArguments (S);
  GetControls;
  //
  GetInfoListe;
  //
  LCHAMPS.Align := alClient;
  LLIBELLES.Align := AlClient;
  LCHPSELECT.Align := alClient;
  LLIBSELECT.align := Alclient;
  //
  LaListeChps := TlistChps.Create;
  RemplitLaliste;
  SetListChamps;
  //
  SetInfoListeRecup;
  //
  SetShow;
  SetEvents(True);
  if not fNewListe then
  begin
    if LCHPSELECT.Items.Count > 0 then
    begin
      LCHPSELECT.ItemIndex := 0;
      CHPSELClick (self);
      THPanel(GetControl('PINFOS')).Enabled := false;
    end;
  end else
  begin
    THPanel(GetControl('PINFOS')).Enabled := true;
  end;
  fDataChanged := false;
  fDataModif := false;
end ;

procedure TOF_BTCRESAISIE.OnClose ;
begin
  if fDataChanged then
  begin
    if PgiAsk ('ATTENTION : Vous allez perdre les modifications.#13#10 Confirmez-vous l''abandon ?') <> mrYes then
    begin
      ecran.ModalResult := 0;
      exit;
    end;
  end;
  TS.Free;
  LaListeChps.Clear;
  TOBLISTE.free;
  Inherited ;
end ;

procedure TOF_BTCRESAISIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCRESAISIE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCRESAISIE.RemplitLaliste;
var UneListe,UnRecord,UR,UnChamps,Modif,Libelle,TypeInfo : string;
    Taille : integer;
    UnElt : TChps;
begin
  UneListe := LISTCHPS;
  if VH_GC.BTCODESPECIF = '001' then
  begin
    UneListe := UneListe + ';SUMTOTALTS|N;MTTRANSFERT|N;NUMTRANSFERT|N;'; // ajout colonne total des avenants sur ligne pour POC
  end;
  repeat
    UnRecord := READTOKENST(UneListe);
    if Unrecord <> '' then
    begin
      UR := Unrecord;
      UnChamps := READTOKENPipe(UR,'|');
      Modif := 'O';
      if UR <> '' then Modif := UR;
      //
      GetInfoChamps (Unchamps,TypeInfo,Libelle,Taille);
      //
      UnElt := TChps.Create;
      UnElt.Nomchps := UnChamps;
      UnElt.LibelleINI := Libelle;
      UnElt.TailleINI := Taille;
      UnElt.TypeChps := TypeInfo;
      UnElt.ModifINI := (Modif='O');
      UnElt.Modifiable := UnElt.ModifINI;
      UnElt.Designation := UnElt.LibelleINI;
      UnElt.TailleReele := Taille;
      UnElt.Obligatoire := false;
      UnElt.LibComplet := false;
      UnElt.BlancSiNul := false;
      UnElt.WithCumul := false;
      UnElt.visible := true;
      if TypeInfo ='INTEGER' then
      begin
        UnElt.NbDec := 0;
        UnElt.NbDecIni := 0;
        UnElt.AlignINI := TalRight;
        UnElt.Align := TalRight;
        UnElt.SepMilliers := false;
      end else if TypeInfo = 'DATE' then
      begin
        UnElt.NbDec := 0;
        UnElt.NbDecIni := 0;
        UnElt.AlignINI := TalRight;
        UnElt.Align := TalRight;
        UnElt.SepMilliers := false;
      end else if TypeInfo = 'RATE' then
      begin
        UnElt.NbDec := 2;
        UnElt.NbDecIni := 2;
        UnElt.AlignINI := TalRight;
        UnElt.Align := TalRight;
        UnElt.SepMilliers := true;
      end else if TypeInfo = 'BOOLEAN' then
      begin
        UnElt.NbDec := 0;
        UnElt.NbDecIni := 0;
        UnElt.AlignINI := TalCenter;
        UnElt.Align := TalCenter;
        UnElt.SepMilliers := false;
      end else if TypeInfo = 'COMBO' then
      begin
        UnElt.NbDec := 0;
        UnElt.NbDecIni := 0;
        UnElt.AlignINI := TalLeft;
        UnElt.Align := TalLeft;
        UnElt.SepMilliers := false;
      end else if copy(TypeInfo,1,7) = 'VARCHAR' then
      begin
        UnElt.NbDec := 0;
        UnElt.NbDecIni := 0;
        UnElt.AlignINI := TalLeft;
        UnElt.Align := TalLeft;
        UnElt.SepMilliers := false;
      end;
      LaListeChps.Add(UnElt);
    end;
  until UnRecord = '';
end;

{ Tlistref }

function TlistChps.Add(AObject: TChps): Integer;
begin
	result := inherited add(Aobject);
end;

destructor TlistChps.destroy;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TChps(Items [Indice])<> nil then TChps(Items [Indice]).free;
    end;
  end;
  inherited;
end;

function TlistChps.find(NomChps: string): TChps;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].Nomchps = NomChps then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TlistChps.GetItems(Indice: integer): TChps;
begin
  result := TChps (Inherited Items[Indice]);
end;

procedure TlistChps.SetItems(Indice: integer; const Value: TChps);
begin
  Inherited Items[Indice]:= Value;
end;

procedure TOF_BTCRESAISIE.SetListChamps;
var II : integer;
    CHP  : TChps;
    AA : string;
begin
  LCHAMPS.Clear;
  LLIBELLES.clear;
  for II := 0 to LaListeChps.Count -1 do
  begin
    CHP := LaListeChps.Items[II];
    AA := CHP.Nomchps;
    LCHAMPS.AddItem(AA,CHP);
    AA := CHP.LibelleINI;
    LLIBELLES.AddItem(AA,CHP);
  end;
end;

procedure TOF_BTCRESAISIE.SelectionneThisItem (TheItem: TChps);
var IndBis : integer;
begin
  LCHPSELECT.AddItem(TheItem.Nomchps,TheItem);
  LLIBSELECT.AddItem(TheItem.LibelleINI,TheItem);
  //
  IndBis := LCHAMPS.Items.IndexOf(TheItem.Nomchps);
  LCHAMPS.Items.Delete(IndBis);
  //
  IndBis := LLIBELLES.Items.IndexOf(TheItem.LibelleINI);
  LLIBELLES.Items.Delete(IndBis);
  //
  if RBCHPS.Checked then
  begin
    LCHPSELECT.ItemIndex := LCHPSELECT.Items.count-1;
    CHPSELClick (self);
  end else
  begin
    LLIBSELECT.ItemIndex := LLIBSELECT.Items.count-1;
    LIBSELClick (self);
  end;
end;

procedure TOF_BTCRESAISIE.BInClick(Sender: Tobject);
var TheItem : TChps;
begin
  TheItem := nil;
  if fIndexItem >= 0 then
  begin
    EnregInfo;
  end;
  if RBCHPS.Checked then
  begin
    fIndexItem := LCHAMPS.ItemIndex;
    if fIndexItem >= 0 then TheItem := TChps(LCHAMPS.Items.Objects[LCHAMPS.ItemIndex]);
  end else
  begin
    fIndexItem := LLIBELLES.ItemIndex;
    if fIndexItem >= 0 then TheItem := TChps(LLIBELLES.Items.Objects[LLIBELLES.ItemIndex]);
  end;
  if TheItem <> nil then
  begin
    SelectionneThisItem (TheItem);
    LCHAMPS.Update;
    LLIBELLES.Update;
    fDataChanged := true;
  end;
end;

procedure TOF_BTCRESAISIE.BOutClick(Sender: Tobject);
var TheItem : TChps;
    AA : string;
    IndBis,NextInd : integer;
begin
  fIndexItem := -1;
  NextInd := -1;
  if RBCHPS.Checked then
  begin
    if LCHPSELECT.ItemIndex = LCHPSELECT.Items.Count -1 then
    begin
      NextInd := LCHPSELECT.Items.Count -2
    end else if (LCHPSELECT.ItemIndex = 0) and (LCHPSELECT.Items.Count >0) then
    begin
      NextInd := 0;
    end else NextInd := LCHPSELECT.ItemIndex;

    TheItem := TChps(LCHPSELECT.Items.Objects[LCHPSELECT.ItemIndex]);
  end else
  begin
    if LLIBSELECT.ItemIndex = LLIBSELECT.Items.Count -1 then
    begin
      NextInd := LLIBSELECT.Items.Count -2
    end else if (LLIBSELECT.ItemIndex = 0) and (LLIBSELECT.Items.Count >0) then
    begin
      NextInd := 0;
    end else NextInd := LLIBSELECT.ItemIndex;
    TheItem := TChps(LLIBSELECT.Items.Objects[LLIBSELECT.ItemIndex]);
  end;
  if TheItem.Obligatoire then
  begin
    PGIInfo('Vous ne pouvez pas enlever une zone obligatoire');
    exit;
  end;
  IndBis := LCHPSELECT.Items.IndexOf(TheItem.Nomchps);
  LCHPSELECT.Items.Delete(IndBis);
  LCHPSELECT.Update;
  //
  IndBis := LLIBSELECT.Items.IndexOf(TheItem.LibelleINI);
  LLIBSELECT.Items.Delete(IndBis);
  LLIBSELECT.Update;
  // reinit
  TheItem.TailleReele := TheItem.TailleINI;
  TheItem.Designation := TheItem.LibelleINI;
  TheItem.NbDEC := TheItem.NBDECINI;
  TheItem.BlancSiNul := false;
  TheItem.AlignINI := TheItem.Align;
  //
  AA := TheItem.Nomchps;
  LCHAMPS.AddItem(AA,TheItem);
  AA := TheItem.LibelleINI;
  LLIBELLES.AddItem(AA,TheItem);
  fDataChanged := true;

  if NextInd >= 0 then
  begin
    if RBCHPS.Checked then
    begin
      LCHPSELECT.ItemIndex := NextInd;
      CHPSELClick (self);
    end else
    begin
      LLIBSELECT.ItemIndex := NextInd;
      LIBSELClick(self);
    end;
  end;
end;

procedure TOF_BTCRESAISIE.SetShow;
begin
  LCHAMPS.Align := alClient;
  LLIBELLES.Align := AlClient;
  LCHPSELECT.Align := alClient;
  LLIBSELECT.align := Alclient;
  if RBCHPS.Checked then
  begin
    LCHAMPS.visible := true;
    LCHPSELECT.visible := true;
    LLIBELLES.visible := false;
    LLIBSELECT.visible := false;
  end else
  begin
    LCHAMPS.visible := false;
    LCHPSELECT.visible := false;
    LLIBELLES.visible := true;
    LLIBSELECT.visible := true;
  end;
end;

procedure TOF_BTCRESAISIE.DecodeArguments (Arg : string);
var Par,PP,P1 : string;
begin
  Par := Arg;
  repeat
    PP:= READTOKENST(Par);
    if PP = '' then break;
    P1 := READTOKENPipe (PP,'=');
    if P1 = 'LISTE' then fListeSaisie := PP;
  until PP='';
end;

procedure TOF_BTCRESAISIE.SetInfoListeRecup;
var NomChps,CurChps : string;
    TailleChps,CurTaill : string;
    LibChps,CurLib : string;
    DefChps,Curdef : string;
    TheItem : TChps;
    IndElt,IndBis : integer;
    LAlig,LSepMil,LnbDec,LLibC,LBlanc,LCUMUL,LLIbvisi,LOblig : string;
begin
  NomChps := TS.Strings[1];
  LibChps := TS.Strings[4];
  TailleChps := TS.Strings[5];
  DefChps := TS.Strings[6];
  Repeat
    CurChps := READTOKENST (NomChps);
    if (CurChps ='') or (CurChps=' ') then break;
    CurLib := READTOKENST (LibChps);
    Curdef := READTOKENST (DefChps);
    CurTaill := READTOKENST (TailleChps);
    // Decodage de la definition
    LAlig := Copy(CurDef,1,1);
    LSepMil := Copy(CurDef,2,1);
    LnbDec := Copy(CurDef,3,1);
    LOblig := Copy(CurDef,4,1);
    LLibC := Copy(CurDef,5,1);
    LLibVisi := Copy(CurDef,6,1);
    LBlanc := Copy(CurDef,7,1);
    LCUMUL :=Copy(CurDef,8,1);
    // ---
    IndElt := LCHAMPS.Items.IndexOf(CurChps);
    TheItem := TChps(LCHAMPS.Items.Objects[Indelt]);
    IndBis := LLIBELLES.Items.IndexOf(TheItem.LibelleINI);
    TheItem.Designation := CurLib;
    TheItem.TailleReele := StrToInt(CurTaill);
    TheItem.NbDec := StrToInt(LNbdec);
    if LAlig = 'G' then TheItem.Align := TalLeft
    else if LAlig = 'D' then TheItem.Align := TalRight
    else TheItem.Align := TalCenter;
    TheItem.SepMilliers := (LSepMil='/');
    TheItem.Obligatoire := (LOblig='O');
    TheItem.visible := (LLibVisi<>'X');
    TheItem.BlancSiNul := (LBlanc='X');
    TheItem.WithCumul := (LCUMUL='X');
    //
    LCHAMPS.Items.Delete(IndElt);
    LLIBELLES.Items.Delete(IndBis);
    //
    LCHPSELECT.AddItem(TheItem.Nomchps,TheItem);
    LLIBSELECT.AddItem(TheItem.LibelleINI,TheItem);
    //
  Until CurChps='';
end;

procedure TOF_BTCRESAISIE.GetInfoListe;
var QQ: Tquery;
begin
  fNewListe := false;
  QQ := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="'+fListeSaisie+'"',true,1,'',true);
  if not QQ.eof then
  begin
    TOBLISTE.SelectDB('',QQ); 
    ZNOMLISTE.Text := QQ.findField('LI_LIBELLE').AsString;
    ZCODELISTE.Text := QQ.findField('LI_LISTE').AsString;
    TS.SetText(PChar(QQ.FindField('LI_DATA').AsString));
  end else
  begin
    TOBLISTE.InitValeurs(false);  
    fNewListe := true;
    fDataChanged := true;
  end;
  ferme(QQ);
end;

procedure TOF_BTCRESAISIE.ChpsSelect(Sender: Tobject);
var IndElt :integer;
    TheItem : TChps;
begin
  TheItem := TChps(LCHAMPS.Items.Objects[LCHAMPS.ItemIndex]);
  IndElt := LLIBELLES.Items.IndexOf(TheItem.LibelleINI);
  LLIBELLES.ItemIndex := IndElt;
end;

procedure TOF_BTCRESAISIE.LibSelect(Sender: Tobject);
var IndElt :integer;
    TheItem : TChps;
begin
  TheItem := TChps(LLIBELLES.Items.Objects[LLIBELLES.ItemIndex]);
  IndElt := LCHAMPS.Items.IndexOf(TheItem.Nomchps);
  LCHAMPS.ItemIndex := IndElt;
end;

procedure TOF_BTCRESAISIE.CHPSELClick(Sender: Tobject);
var IndElt :integer;
    TheItem : TChps;
begin
  if fIndexItem >= 0 then
  begin
    EnregInfo;  
  end;
  TheItem := TChps(LCHPSELECT.Items.Objects[LCHPSELECT.ItemIndex]);
  fIndexItem := LCHPSELECT.ItemIndex;
  IndElt := LLIBSELECT.Items.IndexOf(TheItem.LibelleINI);
  LLIBSELECT.ItemIndex := IndElt;
  AfficheInfos(TheItem);
end;

procedure TOF_BTCRESAISIE.LIBSELClick(Sender: Tobject);
var IndElt :integer;
    TheItem : TChps;
begin
  if fIndexItem >= 0 then
  begin
    EnregInfo;  
  end;
  TheItem := TChps(LLIBSELECT.Items.Objects[LLIBSELECT.ItemIndex]);
  fIndexItem := LLIBSELECT.ItemIndex;
  IndElt := LCHPSELECT.Items.IndexOf(TheItem.Nomchps);
  LCHPSELECT.ItemIndex := IndElt;
  AfficheInfos(TheItem);
end;

procedure TOF_BTCRESAISIE.AfficheInfos(SelItem: TChps);
begin
  ZTITLE.Text := SelItem.Designation;
  ZLARG.Value := SelItem.TailleReele;
  ZLARG.text := InttoStr(SelItem.TailleReele);
  ZNBDEC.Value := SelItem.NbDec;
  ZNBDEC.text := InttoStr(SelItem.NbDec);
  ZSEPMIL.Checked := SelItem.SepMilliers;
  ZVIDESINUL.Checked := SelItem.BlancSiNul;
  if SelItem.Align = TalLeft then BALLEFT.Down := true
  else if SelItem.Align = TalRight then BALRIGHT.down := true
  else BALCENTER.down := true;
  ZVISIBLE.Checked := SelItem.visible;
  ZCUMUL.Checked := SelItem.WithCumul;
  ZMODIFOK.Checked := SelItem.Modifiable;
  ZMODIFOK.Enabled := SelItem.ModifINI; 
  ZOBLIG.Checked := SelItem.Obligatoire;
  //

  if (SelItem.TypeChps = 'DOUBLE') or (SelItem.TypeChps = 'RATE') or (SelItem.TypeChps = 'INTEGER') or (SelItem.TypeChps = 'EXTENDED') then
  begin
    ZCUMUL.visible := true;
    if (SelItem.TypeChps = 'DOUBLE') or (SelItem.TypeChps = 'RATE') or (SelItem.TypeChps = 'EXTENDED') then
    begin
      THLabel(GetCOntrol('LZNBDEC')).visible := true;
      ZNBDEC.visible := true;
      ZSEPMIL.visible := true;
    end else
    begin
      THLabel(GetCOntrol('LZNBDEC')).visible := false;
      ZNBDEC.visible := false;
      ZSEPMIL.visible := false;
    end;
    ZVIDESINUL.visible := true;
  end else
  begin
    THLabel(GetCOntrol('LZNBDEC')).visible := false;
    ZNBDEC.visible := false;
    ZSEPMIL.visible := false;
    ZVIDESINUL.visible := false;
    ZCUMUL.visible := false;
  end;
  //
  fDataModif := false;
end;

procedure TOF_BTCRESAISIE.BDownClick(Sender: Tobject);
var IndSel : integer;
begin
  if RBCHPS.Checked then
  begin
    IndSel := LCHPSELECT.ItemIndex;
    if IndSel = LCHPSELECT.Items.Count then exit;
    LCHPSELECT.Items.Exchange(IndSel,IndSel+1);
    LLIBSELECT.Items.Exchange(IndSel,IndSel+1);
    fDataChanged := true;
  end else
  begin
    IndSel := LLIBSELECT.ItemIndex;
    if IndSel = LCHPSELECT.Items.Count -1 then exit;
    LCHPSELECT.Items.Exchange(IndSel,IndSel+1);
    LLIBSELECT.Items.Exchange(IndSel,IndSel+1);
    fDataChanged := true;
  end;

end;

procedure TOF_BTCRESAISIE.BUpClick(Sender: Tobject);
var IndSel : integer;
begin
  if RBCHPS.Checked then
  begin
    IndSel := LCHPSELECT.ItemIndex;
    if IndSel > 0 then
    begin
      LCHPSELECT.Items.Exchange(IndSel,IndSel-1);
      LLIBSELECT.Items.Exchange(IndSel,IndSel-1);
      fDataChanged := true;
    end;
  end else
  begin
    IndSel := LLIBSELECT.ItemIndex;
    if IndSel > 0 then
    begin
      LCHPSELECT.Items.Exchange(IndSel,IndSel-1);
      LLIBSELECT.Items.Exchange(IndSel,IndSel-1);
      fDataChanged := true;
    end;
  end;

end;

procedure TOF_BTCRESAISIE.EnregInfo;
var TheItem : TChps;
begin
  if not fDataModif then exit;
  if RBCHPS.Checked then
  begin
    TheItem := TChps(LCHPSELECT.Items.Objects[fIndexItem]);
  end else
  begin
    TheItem := TChps(LLIBSELECT.Items.Objects[fIndexItem]);
  end;
  TheItem.Designation := ZTITLE.Text;
  TheItem.TailleReele := ZLARG.Value;
  TheItem.NbDec := ZNBDEC.Value;
  TheItem.SepMilliers := ZSEPMIL.Checked;
  TheItem.BlancSiNul := ZVIDESINUL.Checked;
  if BALLEFT.Down then TheItem.Align := TalLeft
  else if BALRIGHT.down then TheItem.Align := TalRight
  else TheItem.Align := TalCenter;
  TheItem.visible := ZVISIBLE.Checked;
  TheItem.WithCumul := ZCUMUL.Checked;
  TheItem.Modifiable := ZMODIFOK.Checked;
  TheItem.Obligatoire := ZOBLIG.Checked;
  fDataModif := false;
end;

procedure TOF_BTCRESAISIE.DataChange(Sender: Tobject);
begin
  fDataModif := true;
  fDataChanged := true;
end;

function TOF_BTCRESAISIE.ConstitueLigne(TtypeDataExt : TListeTypeDat) : string;
var II : integer;
    TheItem : TChps;
    STST  : string;
begin
  result := '';
    For II := 0 to LCHPSELECT.items.Count -1 do
    begin
      TheItem := TChps(LCHPSELECT.Items.Objects[II]);
      if TtypeDataExt = TCNomChps then
      begin
        result := result + TheItem.Nomchps +';'
      end else if TtypeDataExt = TClibChps then
      begin
        result := result + TheItem.Designation +';'
      end else if TtypeDataExt = TClargChps then
      begin
        result := result + IntToStr(TheItem.TailleReele)+';'
      end else if TtypeDataExt = TCAligmentChps then
      begin
        STST := '';
        // alignement
        if TheItem.Align = TalLeft then STST := 'G'
        else if TheItem.align = TalRight then STST := 'D'
        else STST := 'C';
        // Separation des milliers
        IF TheItem.SepMilliers then STST:=STST + '/' else STST := STST+'.';
        // Nb de decimales
        STST := STST + InttoStr(TheItem.NbDec);
        // Obligatoire
        if TheITem.Obligatoire then STST := STST + 'O' else STST := STST + ' ';
        // Libellé complet
        if TheItem.LibComplet then  STST := STST + '$' else STST := STST + ' ';
        // Visible
        if TheItem.visible then  STST := STST + '-' else STST := STST + 'X';
        // Blanc si nul
        if TheItem.BlancSiNul then  STST := STST + 'X' else STST := STST + '-';
        // afficher un cumul
        if TheItem.WithCumul then  STST := STST + 'X' else STST := STST + '-';
        //
        result := result + STST+';'
      end else if TtypeDataExt = TCNumChps then
      begin
        result := result + '1;'
      end;
    end;
end;

procedure TOF_BTCRESAISIE.EnregistreListe;
var St : String;
begin
  TS.Clear;
  TS.Add('LIGNE');
  ST := ConstitueLigne(TCNomChps);
  TS.Add(ST);
  TS.Add('');
  TS.Add('');
  ST := ConstitueLigne(TClibChps);
  TS.Add(ST);
  ST := ConstitueLigne(TClargChps);
  TS.Add(ST);
  ST := ConstitueLigne(TCAligmentChps);
  TS.Add(ST);
  TS.Add('');
  ST := ConstitueLigne(TCNumChps);
  TS.Add(ST);
  if fNewListe then
  begin
    TOBLISTE.SetString('LI_LISTE',ZCODELISTE.Text);
    TOBLISTE.SetString('LI_LIBELLE',ZNOMLISTE.Text);
    TOBLISTE.SetString('LI_UTILISATEUR','---');
    TOBLISTE.SetString('LI_SOCIETE','');
    TOBLISTE.SetString('LI_NUMOK','-');
    TOBLISTE.SetString('LI_LANGUE','FRA');
    TOBLISTE.SetString('LI_TRIOK','X');
    TOBLISTE.SetString('LI_CREATEUR',V_PGI.User);
    TOBLISTE.SetString('LI_DOMAINE','B');
    TOBLISTE.SetInteger('LI_NUMVERSION',0);
    TOBLISTE.SetDateTime('LI_DATECREATION',NowH);
  end;
  TOBListe.setString('LI_DATA',TS.GetText);
  TOBLISTE.SetDateTime('LI_DATEMODIF',NowH);
  if fNewListe then
  begin
    TOBLISte.SetAllModifie(true);
    TOBLISTE.InsertDB(nil);
  end else
  begin
    TOBLISTE.UpdateDB;
  end;
  fDataChanged := false;
end;

procedure TOF_BTCRESAISIE.InitEcran;
begin
  SetEvents(false);
  ZTITLE.text := '';
  ZLARG.Text := '';
  ZNBDEC.Text := ''; ZNBDEC.visible := false;
  ZSEPMIL.Checked := false; ZSEPMIL.visible := false;
  ZVIDESINUL.Checked := false; ZVIDESINUL.visible := false;
  BALLEFT.Click;
  ZVISIBLE.Checked := false;
  ZCUMUL.Checked := false;
  ZMODIFOK.Checked := false;
  ZOBLIG.Checked := false;
  THPanel(GetControl('PINFOS')).Enabled := true;
  ZCODELISTE.text:= 'Z';
  ZNOMLISTE.text:= '';
  SetEvents(true);
end;

Initialization
  registerclasses ( [ TOF_BTCRESAISIE ] ) ;
end.
