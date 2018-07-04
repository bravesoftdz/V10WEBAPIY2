{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/09/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAIREGLTCOTRAIT ()
Mots clefs ... : TOF;BTSAIREGLTCOTRAIT
*****************************************************************}
Unit BTSAIREGLTCOTRAIT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
     Messages,
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
     HMsgBox,
     EntGC,
     HSysMenu,
     paramsoc,
     HTB97,
     graphics,
     grids,
     utofAfBaseCodeAffaire,
     Vierge,
     SAISUTIL,
     UTOF ;

Type
  TOF_BTSAIREGLTCOTRAIT = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    stcell : string;
    FF : string;
    DEv : Rdevise;
    Htitre        : THLabel;
    BValider      : TToolbarButton97;
    BExit         : TToolbarButton97;
    //
    TOBSSTrait    : TOB;
    TobTiers      : Tob;
    TobPiece      : Tob;
    TobPieceTrait : Tob;
    TOBPieceRG    : Tob;
    TOBBasesRg    : Tob;
    TOBMM         : TOB;
    TOBAcomptes   : TOB;
    TOBPorcs : TOB;
    //
    LesColonnes   : string;
    //
    GS            : THGRID;
    ColMandataire : Integer;
    ColLibelle    : Integer;
    ColMtFac      : Integer;
    ColMtFrais    : Integer;
    ColMtReglable : Integer;
    ColMtRegle    : Integer;
    //
    MtMarche  : Double;
    TotFrais  : Double;
    TotRegle  : Double;
    //
    Sortir    : Boolean;
    //
    procedure AfficheLaLigne (Arow : integer);
    //
    function GetTOBLigneRegl (Ligne : Integer) : TOB;
    procedure SetInfosEcran;
    procedure ConstitueGrille;
    procedure ChargeLigGrille;
    procedure ChargeTobPieceTrait;
    procedure ControleChamp(Champ, Valeur: String);
    //
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSGetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

    procedure OnValideSaisie(Sender: TObject);
    procedure OnExitSaisie(Sender: TObject);

    Procedure ZoneSuivanteOuOk(var ACol, ARow: integer;var Cancel: boolean);
    Function  ZoneAccessible(var ACol,ARow: integer): boolean;
    procedure RECALCREGLCLICK (Sender : Tobject);
    procedure AffichelaGrille;
    function FindEntreprise(TobPieceTrait: TOB): TOB;


  public
    Action: TActionFiche;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  end ;

Implementation

uses  TntGrids,
      Factutil,
      FactRG,
      BTStructChampSup,
      UCotraitance;

procedure TOF_BTSAIREGLTCOTRAIT.OnNew ;
begin
  Inherited ;
end ;
                            
procedure TOF_BTSAIREGLTCOTRAIT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAIREGLTCOTRAIT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSAIREGLTCOTRAIT.OnLoad ;
begin

  inherited ;


end ;

procedure TOF_BTSAIREGLTCOTRAIT.AffichelaGrille;
var I : integer;
begin
  for I := 0 to TOBpieceTrait.detail.Count - 1 do
  begin
    Affichelaligne (I+1);
  end;
end;

procedure TOF_BTSAIREGLTCOTRAIT.ChargeLigGrille;
begin

  Sortir := false;

//  GS.VidePile(false);

  GS.rowcount := TobPieceTrait.Detail.Count+1;
  GS.DefaultRowHeight := 18;

  if ColMtRegle >= 0 Then gs.cells[ColMtRegle, 0]       := 'Mt Souhaité';
  if ColMtReglable >= 0 Then GS.cells[ColMtReglable, 0] := 'Mt Réglable';
  if ColMtFrais >= 0 Then GS.cells[ColMtFrais, 0]       := 'Mt Frais';
  if ColMtFac >= 0 Then GS.cells[ColMtFac, 0]           := 'Mt Facturé';
  if ColLibelle >= 0 Then GS.cells[ColLibelle, 0]       := 'Nom';
  if ColMandataire >= 0 Then GS.cells[ColMandataire, 0] := 'Cotraitant';

	TFVierge (Ecran).HMTrad.ResizeGridColumns(GS);

end;

procedure TOF_BTSAIREGLTCOTRAIT.OnArgument (S : String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  fMulDeTraitement := true;
  Inherited ;

  Critere := S;

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

  FF:='#';
  if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'0';
       end;
    FF:=FF+'0';
   end;
  //
  TOBPieceRG := TOB(LaTOB.Data);
  TobPieceTrait := TOB(TOBPieceRG.data);
  TOBBasesRg := TOB(TobPieceTrait.data);
  TOBSSTrait := TOB(TOBBasesRg.Data);
  TOBAcomptes := TOB(TOBSSTrait.Data);
  TOBPorcs := TOB(TOBAcomptes.data);
  //
  TobTiers      := LaTob.Detail[0];
  TobPiece      := LaTob.Detail[0].Detail[0];
  //
  //
  if TobTiers.Detail.Count = 0 then
  begin
    PostMessage(Ecran.Handle, WM_CLOSE, 0, 0);
    Exit;
  end;
  //
  if TobPieceTrait.detail.count = 0 then
  begin
    PostMessage(Ecran.Handle, WM_CLOSE, 0, 0);
    exit;
  end;
  //
  if Assigned(GetControl('HTITRE')) then  Htitre := THLabel(GetControl('HTITRE'));
  //
  if Assigned(GetControl('BValider')) then
  begin
    BValider := TToolbarButton97(ecran.FindComponent('BValider'));
    BValider.OnClick := OnValideSaisie;
  end;
  //
  if Assigned(GetControl('BFerme')) then
  begin
    BExit := TToolbarButton97(ecran.FindComponent('BFerme'));
    BExit.OnClick := OnExitSaisie;
  end;

  TToolbarButton97(getcontrol('BTRECALCREGL')).OnClick := RECALCREGLCLICK;
  DEV.Code := TobPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  //
  ChargeTobPieceTrait;
  //
  SetInfosEcran;
  //
  ConstitueGrille;
  //
  ChargeLigGrille;
  AffichelaGrille;
	//
  GS.SetFocus;
  GS.col := ColMtRegle;
  GS.row := 2;
  stCell := GS.cells [GS.col,GS.row];
  //
  GS.OnCellEnter  := GSCellEnter;
  GS.OnRowEnter   := GSRowEnter;
  GS.OnCellExit   := GSCellExit;
  GS.OnRowExit    := GSRowExit;
  GS.GetCellCanvas:= GSGetCellCanvas;
end;

procedure TOF_BTSAIREGLTCOTRAIT.SetInfosEcran;
begin

  SetControlText('GP_TIERS',      Tobpiece.getstring('GP_TIERS'));
  SetControlText('TGP_NOMTIERS',  TobTiers.getstring('T_LIBELLE'));
  Htitre.Caption := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_LIBELLE');
  SetControlText('GP_AFFAIRE0',   Tobpiece.getstring('GP_AFFAIRE0'));
  SetControlText('GP_AFFAIRE1',   Tobpiece.getstring('GP_AFFAIRE1'));
  SetControlText('GP_AFFAIRE2',   Tobpiece.getstring('GP_AFFAIRE2'));
  SetControlText('GP_AFFAIRE3',   Tobpiece.getstring('GP_AFFAIRE3'));
  SetControlText('GP_AVENANT',    Tobpiece.getstring('GP_AVENANT'));
  SetControlText('GP_DATEPIECE',  Tobpiece.getstring('GP_DATEPIECE'));
  SetControlText('GP_NUMERO',     Tobpiece.getstring('GP_NUMERO'));
  SetControlText('GP_ETABLISSEMENT', Tobpiece.getstring('GP_ETABLISSEMENT'));
  SetControlText('GP_REFINTERNE', Tobpiece.getstring('GP_REFINTERNE'));
  SetControlText('GP_REFEXTERNE', Tobpiece.getstring('GP_REFEXTERNE'));

end;
procedure TOF_BTSAIREGLTCOTRAIT.ConstitueGrille;
var NbCol : Integer;
    st    : string;
    Nam   : String;
    i     : Integer;
    depart: Integer;
begin

  if not Assigned(GetControl('GS_FRAIS')) then exit;
  //
  GS := THGRID(GetControl('GS_FRAIS'));
  //
  NbCol := 8;
  //
//  LesColonnes := 'FIXED;BPE_TYPEINTERV;BPE_FOURNISSEUR;NOMFRS;MONTANTFACT;TOTALFRAIS;MONTANTREGLABLE;BPE_MONTANTREGL';
  LesColonnes := 'FIXED;BPE_TYPEINTERV;BPE_FOURNISSEUR;LIBELLE;MONTANTFACT;TOTALFRAIS;MONTANTREGLABLE;BPE_MONTANTREGL';
  //
  GS.ColCount     := NbCol;
  GS.ColWidths[0] := 10;
  //
  St := LesColonnes;
  //
  ColMandataire   := -1;
  ColLibelle      := -1;
  ColMtFac        := -1;
  ColMtFrais      :=  1;
  ColMtReglable   := -1;
  ColMtRegle      := -1;
  //
  i := 0;
  repeat
    Nam := ReadTokenSt(St);
    if Nam = '' then Break;
    //
    if Nam = 'FIXED' then
    begin
      GS.ColWidths[i] := 20;
      GS.ColEditables[i] := False;
    end else if Nam = 'BPE_FOURNISSEUR' then
    begin
      GS.ColWidths[i] := 60;
      GS.ColEditables[i] := False;
      ColMandataire := i;
    end else if Nam = 'BPE_TYPEINTERV' then
    begin
      GS.ColWidths[i] := 20;
      GS.ColEditables[i] := False;
    	GS.ColFormats[i] := 'CB=AFINTERVENANTAFFAIRE';
      GS.ColDrawingModes[I]:= 'IMAGE';
    end else if Nam = 'LIBELLE' then
    begin
      GS.ColWidths[i] := 300;
      GS.ColEditables[i] := False;
      ColLibelle := i;
    end else if Nam = 'MONTANTFACT' then
    begin
      GS.ColWidths[i] := 80;
      GS.ColFormats[i] := FF;
      GS.ColAligns[i] := taRightJustify;
      GS.ColEditables[i] := False;
      ColMtFac := i;
    end else if (Nam = 'TOTALFRAIS') then
    begin
      GS.ColWidths[i] := 80;
      GS.ColFormats[i] := FF;
      GS.ColAligns[i] := taRightJustify;
      GS.ColEditables[i] := False;
      ColMtFrais := i;
    end else if (Nam = 'MONTANTREGLABLE') then
    begin
      GS.ColWidths[i] := 80;
      GS.ColFormats[i] := FF;
      GS.ColAligns[i] := taRightJustify;
      ColMtReglable := i;
    end else if (Nam = 'BPE_MONTANTREGL') then
    begin
      GS.ColWidths[i] := 90;
      GS.ColFormats[i] := FF;
      GS.ColAligns[i] := taRightJustify;
      ColMtRegle := i;
    end;
    inc(i);
  until nam = '';

end ;

procedure TOF_BTSAIREGLTCOTRAIT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSAIREGLTCOTRAIT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAIREGLTCOTRAIT.OnCancel () ;
begin
  Inherited ;
end ;

function Tof_BTSAIREGLTCOTRAIT.FindEntreprise (TobPieceTrait : TOB) : TOB;
begin
  result := TOBPieceTrait.findFirst(['BPE_FOURNISSEUR'],[''],True);
end;

procedure Tof_BTSAIREGLTCOTRAIT.ChargeTobPieceTrait;
var CodeFrs   : String;
    Affaire   : String;
    iInd      : integer;
    TOBMANDAT : TOB;
begin

  if TobPieceTrait = nil then Exit;
  //
//  TOBPieceRG := TOB.Create ('LES RETENUES',nil,-1);
//  TOBBasesRg := TOB.Create ('LES BASESRG',nil,-1);
  //
  Affaire    := TobPIECE.GetString('GP_AFFAIRE');
  //
  MtMarche   := TobPiece.GetDouble('GP_TOTALTTCDEV') - TobPiece.GetDouble('GP_ACOMPTEDEV');
  TotFrais   := 0;
  TotRegle   := 0;

  // LoadLesRetenues(TOBPiece, TOBPieceRG, TOBBasesRG,taModif);
  CalculeReglementsIntervenants (TOBSSTrait,TobPiece,TOBPieceRG,TOBAcomptes,TOBporcs,Affaire,TobPieceTrait,DEv,false);

  TOBMM := FindEntreprise (TobPieceTrait);

end;


procedure TOF_BTSAIREGLTCOTRAIT.GSCellEnter(Sender: TObject; var ACol,  ARow: Integer; var Cancel: Boolean);
begin

  if Action = taConsult then Exit;
  stCell := GS.cells [Acol,Arow];
  ZoneSuivanteOuOk(ACol, ARow, Cancel);

end;

procedure TOF_BTSAIREGLTCOTRAIT.GSCellExit(Sender: TObject; var ACol,  ARow: Integer; var Cancel: Boolean);
var Affaire : string;
    CodeFrs : string;
    MtCalcul    : Double;
    MtMandat    : Double;
    MtReglable  : Double;
    MtSouhait   : Double;
    TOBL : TOB;
begin

  if GS.cells[Acol,Arow]= stcell then Exit;
  //
  TOBL := GetTOBLigneRegl(Arow);
  //
  MtReglable  := TOBL.GetDouble('BPE_MONTANTREGL');
  MtMandat    := TOBMM.GetDouble('BPE_MONTANTREGL'); // Ligne entreprise
  MtSouhait   := Valeur(GS.Cells[ColMtRegle, ARow]);


  if (MtSouhait - MtReglable) > MtMandat then
  begin
    PGIError('La différence entre souhaité et reglable du cotraitant est supérieure au montant réglable du mandataire', 'erreur saisie');
    MtSouhait := MtReglable;
    cancel := true;
  end;

  Affaire     := TobPiece.GetString('GP_AFFAIRE');
  CodeFrs     := TOBL.GetString('BPE_FOURNISSEUR');

  MtCalcul    := MtMandat + MtReglable - MtSouhait;

  TOBL.SetBoolean('BPE_REGLSAISIE',true);
  TOBL.PutValue('BPE_MONTANTREGL', MtCalcul);
  TOBL.PutValue('BPE_MONTANTREGL', MtSouhait);
  TOBMM.SetDouble('BPE_MONTANTREGL',TOBMM.GetDouble('BPE_MONTANTREGL')-MtSouhait + MtReglable);
  Affichelaligne(1);
  Affichelaligne(Arow);
end;

procedure TOF_BTSAIREGLTCOTRAIT.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BTSAIREGLTCOTRAIT.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin

end;

function TOF_BTSAIREGLTCOTRAIT.ZoneAccessible(var ACol, ARow: integer): boolean;
begin

  Result := False;

  if (Acol = ColMtRegle) and (ARow <> 1) then
  begin
    Result := True;
    Exit;
  end;

end;

procedure TOF_BTSAIREGLTCOTRAIT.ZoneSuivanteOuOk(var ACol, ARow: integer; var Cancel: boolean);
var Sens    : Integer;
    ii      : Integer;
    Lim     : Integer;
    OldEna  : Boolean;
    ChgLig  : Boolean;
    ChgSens : Boolean;
begin

  OldEna := GS.SynEnabled;
  GS.SynEnabled := False;

  ii := 0;

  Sens := -1;

  ChgLig := (GS.Row <> ARow);
  ChgSens := False;

  if GS.Row > ARow then Sens := 1 else if ((GS.Row = ARow) and (ACol <= GS.Col)) then Sens := 1;

  ACol := GS.Col;
  ARow := GS.Row;

  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GS.RowCount ;
      // ---
      if ((ACol = GS.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          // Ajout d'une ligne
          break;
        end;
      end;
      if ChgLig then
      begin
        ACol := GS.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GS.ColCount - 1 then
      begin
        Inc(ACol);
      end else
      begin
        Inc(ARow);
        ACol := GS.FixedCols;
      end;
    end else
    begin
      if ((ACol = GS.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := GS.ColCount;
        ChgLig := False;
      end;
      if ACol > GS.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GS.ColCount - 1;
      end;
    end;
  end;

  GS.SynEnabled := OldEna;

end;

procedure TOF_BTSAIREGLTCOTRAIT.OnValideSaisie(Sender : TObject);
var iInd : integer;
		TOBP : TOB;
    Arow,Acol : integer;
    cancel : boolean;
begin
  Arow := GS.row;
  Acol := GS.col;
  GSCellExit (self,Acol,Arow,cancel);
  if cancel then exit;
	iInd := 0;
  repeat
    TOBP := TOBPieceTrait.Detail[iInd];
    if (TOBP.getvalue('BPE_TOTALHTDEV')=0) and (TOBP.getvalue('BPE_MONTANTREGL')=0) then
    begin
    	TOBP.free;
    end else
    begin
    	inc(iInd);
    end;
  until iInd >= TOBPieceTrait.Detail.Count;

  //For iInd := 0 to tobpieceTrait.detail.count -1 do
  //begin
    TobPieceTrait.SetAllModifie(True);
    TobPieceTrait.InsertOrUpdateDB;
    LaTOB.PutValue('VALIDEOK','X');
  //end;

end;

procedure TOF_BTSAIREGLTCOTRAIT.OnExitSaisie(Sender: TObject);
begin

end;

procedure TOF_BTSAIREGLTCOTRAIT.Affichelaligne(Arow: integer);
begin
  TobPieceTrait.Detail[Arow-1].PutLigneGrid(GS,Arow,False,False,lesColonnes);
end;

procedure TOF_BTSAIREGLTCOTRAIT.GSGetCellCanvas(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
  if ACol < GS.FixedCols then Exit;
  IF Arow < GS.fixedRows then exit;
  //
  if (Arow = 1) then BEGIN canvas.Brush.Color := clInfoBk; END;
  //

end;

procedure TOF_BTSAIREGLTCOTRAIT.NomsChampsAffaire(var Aff, Aff0, Aff1,
  Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  Aff  := THEdit(GetControl('GP_AFFAIRE'));
  Aff1 := THEdit(GetControl('GP_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GP_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GP_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GP_AVENANT'));
end;

procedure TOF_BTSAIREGLTCOTRAIT.ControleChamp(Champ, Valeur: String);
begin

  if champ = 'COTRAITANCE' then
  begin
    Ecran.Caption := 'Saisie des Réglements Cotraitants';
  end
  else if champ = 'SOUSTRAITANCE' then
  begin
    Ecran.Caption := 'Saisie des Réglements sous-Traitants';
  end;

end;


procedure TOF_BTSAIREGLTCOTRAIT.RECALCREGLCLICK(Sender: Tobject);
var Indice : integer;
begin
  for Indice:= 0 to  TobPieceTrait.detail.count -1 do
  begin
    TobPieceTrait.detail[Indice].SetBoolean('BPE_REGLSAISIE',false);
    TobPieceTrait.detail[Indice].SetDouble('BPE_MONTANTREGL',TobPieceTrait.detail[Indice].GetDouble('MONTANTREGLABLE'));
  end;
	ChargeLigGrille;
  AffichelaGrille;
  stCell := GS.cells [GS.col,GS.row];
end;

function TOF_BTSAIREGLTCOTRAIT.GetTOBLigneRegl(Ligne: Integer): TOB;
begin
	Result := TobPieceTrait.detail[Ligne-1];
end;

Initialization
  registerclasses ( [ TOF_BTSAIREGLTCOTRAIT ] ) ;
end.

