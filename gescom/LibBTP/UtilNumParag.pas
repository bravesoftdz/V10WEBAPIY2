unit UtilNumParag;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HEnt1,
     UTOB,
     Ent1,
     SysUtils,
     UtilPGI,
     AGLInit,
     HCtrls,
     forms,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     Maineagl,
     eMul,
{$ENDIF}
     EntGC, Classes, HMsgBox,Menus;

const MAXITEMS = 1;
type

  TNumParag = class
  private
    XX : TForm;
    ITypeLigne,ICompteur,ICompteurBis : integer;
    TOBCompteur : TOB;
    fMaxNumP : integer;
    fIsOneForce : boolean;
    fActif : boolean;
    fTOBPiece : TOB;
    fTOBForced : TOB;
    //
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    POPGS : TPopupMenu;
    fcreatedPop : boolean;
    fMaxItems : integer;
    //
    procedure AddSupParag;
    procedure AjouteNiveau (TOBL : TOB);
    function EncodeRefParag: string;
    procedure RemonteNiveau(TOBL: TOB);
//    procedure MemoriseInfo (TOBL : TOB);
    function GetNumparag(TOBL: TOB): integer;
    function OkMemoriseNiv1(TOBL: TOB): boolean;
    procedure DefiniMenuPop(Parent : Tform);
    procedure AnnulForcedNum (Sender : TObject);
    procedure SetActif(const Value: boolean);
    procedure MemoriseLigne (TOBL : TOB);
  public
    constructor create (FF : Tform);
    destructor destroy ; override;

    property MaxNump : integer read fMaxNump write fMaxNump;
    property IsOneManuel : boolean read fIsOneForce;
    property Actif : boolean read fActif write SetActif ;
    property TOBPiece : TOB read fTOBPiece write fTOBPiece;

    function ChangeCodeParag(TOBpiece: TOB; Arow: integer;Valeur: string) : boolean;
    procedure ConstitueParag (TOBPiece : TOB; Debut,Fin : integer; Force : boolean=false);
    procedure InitCompteur (Niveau : integer);
    procedure MemoriseMax (TOBL : TOB);
    procedure SetInfoLigne (TOBPiece : TOB; Arow : integer);
    procedure ReinitCompteurMax;
    procedure SetPopMenu (TOBL : TOB);
    procedure DeleteLigne (Arow : integer);
    procedure Reinit(TOBpiece: TOB);
  end;


implementation
uses facture;

{ TNumParag }

procedure TNumParag.AddSupParag;
begin
  TOBCompteur.AddChampSupValeur('PN1','');
  TOBCompteur.AddChampSupValeur('PN2','');
  TOBCompteur.AddChampSupValeur('PN3','');
  TOBCompteur.AddChampSupValeur('PN4','');
  TOBCompteur.AddChampSupValeur('PN5','');
  TOBCompteur.AddChampSupValeur('PN6','');
  TOBCompteur.AddChampSupValeur('PN7','');
  TOBCompteur.AddChampSupValeur('PN8','');
  TOBCompteur.AddChampSupValeur('PN9','');
  TOBCompteur.AddChampSupValeur('PN10','');
end;

procedure TNumParag.AjouteNiveau(TOBL: TOB);
var NivP : integer;
begin
  NivP := -1;
  if (iTypeLigne = -1) then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
  if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
  if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');

  if (copy(TOBL.GetString(iTypeLigne),1,2)='DP') or (copy(TOBL.GetString(iTypeLigne),1,2)='DV') then
  begin
    NivP := StrToInt(copy(TOBL.GetString(iTypeLigne),3,1));
  end else if (copy(TOBL.GetString(iTypeLigne),1,2)='AR') then
  begin
    NivP := TOBL.GetInteger(iCompteurBis)+1;
  end;
  if NivP = -1 then exit;
  TOBCompteur.SetString('PN'+InttoStr(NivP),InttoStr(TOBCompteur.getInteger('PN'+InttoStr(NivP))+1));
  TOBL.SetString(iCompteur,EncodeRefParag);
  TFfacture(XX).AfficheLaLigne(TOBL.GetIndex +1);
end;

constructor TNumParag.create (FF : Tform);
var ThePop : Tcomponent;
begin
  fTOBForced := TOB.Create ('LES LIGNES FORCEES',nil,-1);
  TOBCompteur := TOB.create ('LE PARAG',nil,-1);
  AddSupParag ;
  iTypeLigne := -1;
  iCompteur := -1;
  iCompteurBis := -1;
  fIsOneForce := false;
  fMaxNumP := 0;
  factif := false;
  //
  XX := FF;
  ThePop := XX.Findcomponent  ('POPBTP');
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(XX);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(XX);
end;

destructor TNumParag.destroy;
var Indice : integer;
begin
  TOBCompteur.free;
  fTOBForced.free;
  for Indice := 0 to fMaxItems -1 do
    begin
    MesMenuItem[Indice].Free;
    end;
  if fcreatedPop then POPGS.free;
  inherited;
end;

function TNumParag.EncodeRefParag : string;
begin
  Result := TOBcompteur.GetString('PN1');
  if TOBcompteur.GetString('PN2')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN2');
  if TOBcompteur.GetString('PN3')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN3');
  if TOBcompteur.GetString('PN4')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN4');
  if TOBcompteur.GetString('PN5')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN5');
  if TOBcompteur.GetString('PN6')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN6');
  if TOBcompteur.GetString('PN7')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN7');
  if TOBcompteur.GetString('PN8')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN8');
  if TOBcompteur.GetString('PN9')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN9');
  if TOBcompteur.GetString('PN10')<>'' then Result:= Result +'.'+ TOBcompteur.GetString('PN10');
end;

procedure TNumParag.RemonteNiveau (TOBL : TOB);
var NivP : integer;
begin
  NivP := StrToInt(copy(TOBL.GetString(iTypeLigne),3,1));
  InitCompteur (Nivp);
end;

procedure TNumParag.InitCompteur (Niveau : integer);
begin
  if not factif then exit;
  if Niveau = 0 then TOBcompteur.SetString('PN1','');
  if Niveau < 2 then TOBcompteur.SetString('PN2','');
  if Niveau < 3 then TOBcompteur.SetString('PN3','');
  if Niveau < 4 then TOBcompteur.SetString('PN4','');
  if Niveau < 5 then TOBcompteur.SetString('PN5','');
  if Niveau < 6 then TOBcompteur.SetString('PN6','');
  if Niveau < 7 then TOBcompteur.SetString('PN7','');
  if Niveau < 8 then TOBcompteur.SetString('PN8','');
  if Niveau < 9 then TOBcompteur.SetString('PN9','');
  if Niveau < 10 then TOBcompteur.SetString('PN10','');
end;

function TNumParag.ChangeCodeParag(TOBpiece: TOB; Arow : integer; Valeur : string) : boolean;

  function ISValeurCompatible(Valeur,TypeLigne : string; Niveau : integer) : boolean;
  var prof : integer;
      LL,ZZ : string;
  begin
    result := false;
    prof := 0;
    ZZ := Valeur;
    repeat
      LL := READTOKENPipe(ZZ,'.');
      if (LL <> '') and (IsNumeric (LL)) then
      begin
        inc(Prof);
      end;
    until LL = '';
    if ((Copy(TypeLigne,1,2) = 'DP') or ((Copy(TypeLigne,1,2) = 'DV'))) and (Niveau = Prof) then result := true;
    if (Copy(TypeLigne,1,2) = 'AR') and (Niveau = Prof-1) then result := true;
  end;

var TOBL : TOB;
    Niveau : integer;
    TypeLigne : string;
    II,IIFIN : integer;
begin
  result := false;
  if not factif then exit;
  TOBL := TOBPiece.detail[Arow-1];
  IIFIN := -1;
  //
  if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
  if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
  if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
  //
  TypeLigne :=TOBL.GetString(ITypeLigne);
  Niveau := TOBL.GetInteger(iCompteurBis);
//  if ISValeurCompatible(Valeur,TypeLigne,Niveau) then
//  begin
    TOBL.SetString('GLC_NUMEROTATION',Valeur);
    TOBL.SetString('GLC_NUMFORCED','X');
    if TOBL.GetString('GL_PIECEPRECEDENTE')<> '' then exit;
    fIsOneForce := true;
    if (Copy(TypeLIgne,1,2)='DP') or (Copy(TypeLIgne,1,2)='DV') then
    begin
      for II := Arow to TOBpiece.detail.count - 1 do
      begin
        if (TOBPiece.Detail[II].GetString(iTypeLigne) = 'TP'+IntToStr(Niveau)) or (TOBPiece.Detail[II].GetString(iTypeLigne) = 'TV'+IntToStr(Niveau)) then
        begin
          IIFin := II;
          break;
        end;
      end;
      if (IIFin = -1) then exit;
//      if isNumeric(Valeur) then
//      begin
//        MemoriseInfo(TOBPiece.Detail[ARow-1]);
//      end else
//      begin
        InitCompteur (0);
        TOBcompteur.SetString('PN1',valeur);
//      end;
      ConstitueParag (TOBpiece,Arow,IIFin-1);
    end else if Copy(TypeLIgne,1,2)='AR' then
    begin
        if Niveau = 0 then
        begin
          (*
          if isNumeric(Valeur) then
          begin
            MemoriseInfo(TOBL);
          end else
          *)
          begin
            InitCompteur (0);
            TOBcompteur.SetString('PN1',valeur);
          end;
        end;
    end;
    if (TOBcompteur.GetInteger('PN1') > fMaxNump) and (OkMemoriseNiv1(TOBL)) then fMaxNump := TOBcompteur.GetInteger('PN1');
    result := true;
    MemoriseLigne (TOBL);
//  end else
//  begin
//    PgiError('Le code paragraphe n''est pas compatible avec son emplacement');
//  end;
end;

function TNumParag.OkMemoriseNiv1(TOBL: TOB) : boolean;
var Niv : integer;
    TypeLigne : string;
begin
  //
  if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
  if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
  if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
  //
  TypeLigne :=TOBL.GetString(ITypeLigne);
  Niv := TOBL.GetInteger(iCompteurBis);
  result :=  ( ((Copy(TypeLigne,1,2)='DP') or (Copy(TypeLigne,1,2)='DV')) and (Niv=1)) or ((Copy(TypeLigne,1,2)='AR') and (niv=0));
end;


procedure TNumParag.ConstitueParag(TOBPiece: TOB; Debut,Fin : integer; Force : boolean=false);
var indice : integer;
    TOBl : TOB;
    TypeLigne : string;
    ValSaisie : boolean;
begin
  if not factif then exit;
  for indice := Debut to Fin do
  begin
    TOBL := TOBPiece.detail[Indice];
    //
    if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
    if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
    if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
    //
    TypeLigne :=TOBL.GetString(ITypeLigne);
    ValSaisie := TOBL.GetBoolean('GLC_NUMFORCED');
    if (force) or (not valSaisie) then
    begin
      if (fMaxNumP > TOBcompteur.GetInteger('PN1')) and (OkMemoriseNiv1(TOBL)) then TOBcompteur.SetInteger('PN1',fMaxNump);
      //
      if (copy(TypeLigne,1,2)='DP') or (Copy(TypeLigne,1,2) = 'DV') or (Copy(TypeLigne,1,2) = 'AR') then
      begin
        AjouteNiveau (TOBL);
      end else if (copy(TypeLigne,1,2)='TP') or (copy(TypeLigne,1,2)='TV') then
      begin
        RemonteNiveau(TOBL);
      end;
      TOBL.SetBoolean('GLC_NUMFORCED',false);
    end else if (ValSaisie) and ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV') ) then
    begin
      (*
      if IsNumeric (TOBL.GetString('GLC_NUMEROTATION')) then
      begin
        MemoriseInfo(TOBL);
      end else
      *)
      begin
        InitCompteur (0);
        TOBcompteur.SetString('PN1',TOBL.GetString('GLC_NUMEROTATION'));
      end;
      TOBL.SetBoolean('GLC_NUMFORCED',true);
    end;
    //
    if (TOBcompteur.GetInteger('PN1') > fMaxNump) and (OkMemoriseNiv1(TOBL)) then fMaxNump := TOBcompteur.GetInteger('PN1');
  end;
  if Force then fIsOneForce := false;
end;

procedure TNumParag.MemoriseMax(TOBL: TOB);
var NivP,HH : integer;
begin
  if not fActif then exit;
  NIVP := -1;
  //
  if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
  if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
  if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
  //
  if (not fIsOneForce) and (TOBL.GetBoolean('GLC_NUMFORCED')) then fIsOneForce := true;
  if TOBL.GetBoolean('GLC_NUMFORCED') then MemoriseLigne (TOBL);
  if (copy(TOBL.GetString(iTypeLigne),1,2)='DP') or (copy(TOBL.GetString(iTypeLigne),1,2)='DV')then
  begin
    NivP := StrToInt(copy(TOBL.GetString(iTypeLigne),3,1));
  end else if (copy(TOBL.GetString(iTypeLigne),1,2)='AR') then
  begin
    NivP := TOBL.GetInteger(iCompteurBis)+1;
  end;
  if NIVP = -1 then exit;
  if (NIVP = 1) then
  begin
    HH := GetNumparag(TOBL);
    if HH <> 0 then if HH > fMaxNumP then fMaxNumP := HH;
  end;
end;

function TNumParag.GetNumparag(TOBL : TOB) : integer;
var ST : string;
begin
  result := 0;
  if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
  if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
  if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
  //
  St := TOBL.GetString(Icompteur);
  TRY
    if St <> '' then result := StrtoInt(READTOKENPipe(St,'.'));
  EXCEPT
    result := 0;
  END;
end;

procedure TNumParag.SetInfoLigne(TOBPiece: TOB; Arow: integer);
var TOBL : TOB;
    TypeL : string;
    II,IIDep,IIFin,IVal,INiv : integer;
begin
  if not factif then exit;
  TOBL := TOBPiece.Detail[Arow-1];
  IIDep := -1; IIFin := -1;
  //
  if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
  if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
  if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
  //
  IVal := TOBL.getInteger(iCompteurBis);
  TypeL := copy(TOBL.GetString(iTypeLigne),1,2);
  //
  if ((IVal > 1) and ((TypeL='DP') or (TypeL='DV'))) or
     ((IVal >= 1) and (TypeL='AR')) then
  begin
    if (copy(TOBL.GetString(iTypeLigne),1,2)='DP') or (copy(TOBL.GetString(iTypeLigne),1,2)='DV') then iNiv := Ival -1
                                                                                                  else iNiv := Ival;
    for II := Arow-1 downto 0 do
    begin
      if (TOBPiece.Detail[II].GetString(iTypeLigne) = 'DP'+IntToStr(INiv)) or (TOBPiece.Detail[II].GetString(iTypeLigne) = 'DV'+IntToStr(INiv)) then
      begin
        IIdep := II;
        break;
      end;
    end;
    for II := Arow to TOBpiece.detail.count - 1 do
    begin
      if (TOBPiece.Detail[II].GetString(iTypeLigne) = 'TP'+IntToStr(INiv)) or (TOBPiece.Detail[II].GetString(iTypeLigne) = 'TV'+IntToStr(INiv)) then
      begin
        IIFin := II;
        break;
      end;
    end;
    if (IIDep = -1) or (IIFin = -1) then exit;
//    if (TOBPiece.Detail[IIDep].GetString('GLC_NUMFORCED')='X') {or (not Isnumeric(TOBPiece.Detail[IIDep].GetString('GLC_NUMEROTATION')))} then
//    begin
      InitCompteur (0);
      TOBcompteur.SetString('PN1',TOBPiece.Detail[IIDep].GetString('GLC_NUMEROTATION'));
//    end else
//    begin
//      MemoriseInfo(TOBPiece.Detail[IIDep]);
//    end;
    ConstitueParag (TOBpiece,IIDep+1,IIFin-1);
  end else if (pos(TypeL,'DP;AR;DV;')>0) then
  begin
    InitCompteur (0);
    TOBcompteur.PutValue('PN1',fMaxNump);
    ConstitueParag (TOBpiece,Arow-1,Arow-1);
  end;
end;
(*
procedure TNumParag.MemoriseInfo(TOBL: TOB);
var II : integer;
    LL,ZZ : string;
begin
  InitCompteur (0);
  //
  if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
  if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
  if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
  //
  II := 1;
  LL := TOBL.GetString(Icompteur);
  repeat
    ZZ := READTOKENPipe(LL,'.');
    if ZZ <> '' then
    begin
      TOBcompteur.SetInteger('PN'+IntToStr(II),StrToInt(ZZ));
      inc(II);
    end;
  until ZZ = '';
end;
*)
procedure TNumParag.ReinitCompteurMax;
begin
  if not factif then exit;
  fMaxNump := 0;
end;

procedure TNumParag.DefiniMenuPop(Parent: Tform);
var Indice : integer;
begin
  fMaxItems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with MesMenuItem[fMaxItems] do
      begin
      Caption := '-';
      end;
    inc (fMaxItems);
  end;
  // Annulation de la nuémrotation forcée
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'TAnnulNumForced';
    Caption := TraduireMemoire ('Annuler la numérotation forcée');
    OnClick := AnnulForcedNum;
    end;
  MesMenuItem[fMaxItems].ShortCut := ShortCut( Word('A'), [ssCtrl,ssAlt]);
  inc (fMaxItems);
  for Indice := 0 to fMaxItems -1 do
  begin
    if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
  end;
end;

procedure TNumParag.SetPopMenu(TOBL: TOB);
var TheMenuItem : TMenuItem;
  Indice : Integer;
begin
  if not factif then exit;
  TheMenuItem := nil;
  for Indice := 0 to fMaxItems -1 do
  begin
    if (MesMenuItem[Indice].name = 'TAnnulNumForced') then
    begin
      TheMenuItem := MesMenuItem[Indice];
      break;
    end;
  end;

  TheMenuItem.enabled := false;
  if (TheMenuItem <> nil) and (TOBL <> nil)  then
  begin
    if TOBL.GetBoolean('GLC_NUMFORCED') then TheMenuItem.enabled := true;
  end;

end;

procedure TNumParag.AnnulForcedNum(Sender: TObject);
var TOBL,TT : TOB;
begin
  TOBL := fTOBPiece.detail[TFFACTURE(XX).GS.Row-1];
  if TOBL = nil then exit;
  TOBL.SetBoolean('GLC_NUMFORCED',false);
  TT := fTOBForced.FindFirst(['NUMORDRE'],[TOBL.GEtInteger('GL_NUMORDRE')],true);
  if TT <> nil then TT.free;
  fIsOneForce := (fTOBForced.detail.count > 0);
  TFFACTURE(XX).GS.InvalidateRow(TFFACTURE(XX).GS.Row);
end;

procedure TNumParag.SetActif(const Value: boolean);
var Indice : integer;
begin
  fActif := Value;
  if not fActif then
  begin
    for Indice := 0 to fMaxItems -1 do
    begin
      MesMenuItem[Indice].visible := false;
    end;
  end;
end;

procedure TNumParag.MemoriseLigne(TOBL: TOB);
var TT : TOB;
begin
  TT := TOB.Create ('UNE LIGNE',fTOBForced,-1);
  TT.AddChampSupValeur('NUMORDRE',TOBL.GetInteger('GL_NUMORDRE'));
end;

procedure TNumParag.DeleteLigne(Arow: integer);
var TOBL,TT : TOB;
begin
  if not fActif then exit;
  if  fTOBPiece = nil then exit;
  TOBL := fTOBPiece.detail[TFFACTURE(XX).GS.Row-1];
  if TOBL = nil then exit;
  TT := fTOBForced.FindFirst(['NUMORDRE'],[TOBL.GEtInteger('GL_NUMORDRE')],true);
  if TT <> nil then TT.free;
  fIsOneForce := (fTOBForced.detail.count > 0);
end;

procedure TNumParag.Reinit(TOBpiece : TOB);
var indice : integer;
    TOBl : TOB;
    TypeLigne : string;
    ValSaisie : boolean;
begin
  if not factif then exit;
  for indice := 0 to TOBpiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    //
    if (iTypeLigne = -1)  then ItypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
    if (iCompteur = -1) then iCompteur := TOBL.GetNumChamp('GLC_NUMEROTATION');
    if (iCompteurBis = -1) then iCompteurBis := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
    if TOBL.GetString(iCompteur) = '' then continue;
    //
    TOBL.SetBoolean('GLC_NUMFORCED',false);
    TOBL.SetString(iCompteur,'');
  end;
end;

end.
