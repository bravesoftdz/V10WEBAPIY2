{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/10/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPIECEINTERVMUL ()
Mots clefs ... : TOF;BTPIECEINTERVMUL
*****************************************************************}
Unit BTPIECEINTERVMUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
{$ENDIF}
		 AGLInit,
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     Vierge,
     UTOF ;

Type
  TOF_BTPIECEINTERVMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    GS : THGrid;
  	fColNames,Falignement,Ftitre,fLargeur,fAffaire :string;
    BMODIFIER : TToolbarButton97;
    BInsert : TToolbarButton97;
    BDelete : TToolbarButton97;
    TOBParam  : TOB;
    TOBSSTRAIT : TOB;
    TOBPieceTrait : TOB;
    procedure GSDblclick (Sender : TObject);
    procedure OnInsert(Sender: Tobject);
    procedure BmodifierClick (Sender : TObject);
    procedure AfficheLaGrille;
    procedure AfficheLaLigne (Indice : Integer);
    procedure DefiniGrille(Grid: Thgrid; ColNames, Aligments, Titles,Widths: string);
    procedure BdeleteClick (Sender : Tobject);
    function AssignedSSTrait (TOBT : TOB) : boolean;
  end ;

Implementation
uses UCotraitance;

procedure TOF_BTPIECEINTERVMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERVMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERVMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERVMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERVMUL.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  TOBParam := LaTOB;
  TOBSSTRAIT := TOB(LaTOB.Data);
  TOBPieceTrait := TOB(TOBSSTRAIT.data);
  TOBSSTRAIT.data := nil;
  fAffaire := TOBParam.GetValue('AFFAIRE');
  //
  GS := ThGrid(GetControl('GS'));
  BInsert := TToolbarButton97 (GetControl('Binsert'));
  BDelete := TToolbarButton97 (GetControl('BDelete'));
  BMODIFIER := TToolbarButton97 (GetControl('BMODIFIER'));
  //
  fColNames := 'SEL;BPI_TYPEINTERV;BPI_TIERSFOU;TLIBTIERSFOU;BPI_TYPEPAIE;';
  Falignement := 'G.0  ---;C.0O --X;G.0  ---;G.0  ---;C.0O --X;';
  Ftitre := 'Num;Type;Code;Désignation;Paie;';
  fLargeur := '5;5;18;60;5;';
  //
  Definigrille (GS,fColNames,Falignement,Ftitre,fLargeur);
  AfficheLaGrille;
  if TOBParam.GetString('ACTION')='SELECTION' then GS.OnDblClick := GSDblclick;
  BInsert.OnClick := OnInsert;
  BMODIFIER.OnClick := BmodifierClick;
  BDelete.OnClick := BdeleteClick;
  //
  //FV1 - 17/01/2018 : FS#2798 - TESTS BL : en affectation ST dans devis, Erreur si clic sur la loupe et pas de ST
  iF TOBSSTRAIT.Detail.count = 0 THEN
  Begin
    BMODIFIER.Visible := False;
    BDelete.Visible   := False;
  end;

end ;

procedure TOF_BTPIECEINTERVMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERVMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEINTERVMUL.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BTPIECEINTERVMUL.GSDblclick(Sender: TObject);
var Tiers,ModePaie : string;
begin
  Tiers:= TOBSSTRAIT.detail[GS.row-1].GetValue('BPI_TIERSFOU');
  ModePaie:=TOBSSTRAIT.detail[GS.row-1].GetValue('BPI_TYPEPAIE');
  TOBParam.putValue('SOUSTRAIT',Tiers);
  TOBParam.putValue('MODEPAIE',ModePaie);
  Ecran.Close;
end;

procedure TOF_BTPIECEINTERVMUL.OnInsert(Sender: Tobject);
var TTOB : TOB;
begin
  TTOB := TOB.Create('PIECEINTERV',TOBSSTRAIT,-1);
  AddChampsSupTraitCreat (TTOB);
  AddChampsSupTrait (TTOB);
  //
  TheTOB := TTOB;
  TheTOB.Data := TOBSSTRAIT;
  AGLLanceFiche('BTP','BTPIECEINTERV','','','ACTION=CREATION;AFFAIRE='+fAffaire);
  TTOB.Data := nil;
  TheTOB := nil;
  if TTOB.GetValue('BPI_TIERSFOU')='' then TTOB.Free;
  AfficheLaGrille;
  //
  //FV1 - 17/01/2018 : FS#2798 - TESTS BL : en affectation ST dans devis, Erreur si clic sur la loupe et pas de ST
  iF TOBSSTRAIT.Detail.count <> 0 THEN
  Begin
    BMODIFIER.Visible := True;
    BDelete.Visible   := True;
  end;

end;

procedure TOF_BTPIECEINTERVMUL.AffichelaGrille;
var Indice : integer;
begin

  GS.RowCount := TOBSSTRAIT.detail.Count +1;

  if GS.RowCount < 2 then
  begin
    GS.RowCount := 2;
  end;

	for Indice := 0 to TOBSSTRAIT.detail.Count -1 do
  begin
  	AfficheLaLigne (Indice);
  end;

  GS.row := 1;

  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);

end;


procedure TOF_BTPIECEINTERVMUL.DefiniGrille (Grid : Thgrid; ColNames,Aligments, Titles, Widths: string);
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,lelement : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec,NbCols : integer;
    indice,ind : integer;
    FFQTE,FFDEC : string;
begin
  //
  st := ColNames;
  NbCols := 0;
  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  Grid.ColCount := Nbcols;
  //
  st := ColNames;
  lesalignements := Aligments;
  lestitres := Titles;
  leslargeurs := Widths;

  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    grid.cells[Indice,0] := leTitre;
    grid.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G' then grid.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then grid.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then grid.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if OkVisu then
    begin
  		grid.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
    end else
    begin
    	grid.ColWidths[indice] := -1;
    end;

    if (OkLib) or (okImg) then
    begin
    	grid.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	grid.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end else if (Dec<>0) or (Sep) then
    begin
      FFQTE := '#';
      if Dec > 0 then
      begin
        FFQTE := FFQTE+'0.';
        for ind := 1 to Dec do
        begin
          FFQTE := FFQTE + '0';
        end;
      end else FFQTE := FFQTE + '0';
    	if OkNulle then
      begin
        //
				FFDEC := FFQTE+';-'+FFQTE+'; ;';
      	GS.ColFormats[indice] := FFDEC ;
      end else
      begin
      	GS.ColFormats[indice] := FFQTE ;
      end;
    end;
  end ;
end;

procedure TOF_BTPIECEINTERVMUL.AfficheLaLigne(Indice: Integer);
begin
	TOBSSTRAIT.detail[Indice].PutLigneGrid(GS,Indice+1,false,false,fColNames);
end;

procedure TOF_BTPIECEINTERVMUL.BmodifierClick(Sender: TObject);
begin

  TheTOB := TOBSSTRAIT.detail[GS.row-1];
  TheTOB.Data := TOBSSTRAIT;
  AGLLanceFiche('BTP','BTPIECEINTERV','','','ACTION=MODIFICATION');
  TheTOB := nil;
  TOBSSTRAIT.Data := nil;

  AfficheLaGrille;

end;

procedure TOF_BTPIECEINTERVMUL.BdeleteClick(Sender: Tobject);
begin
	if not AssignedSSTrait (TOBSSTRAIT.detail[GS.row-1]) then
  begin
    if PGIAsk('Etes-vous sur de vouloir supprimer ce sous-traitant ?') = mryes then
    begin
      TOBSSTRAIT.detail[GS.row-1].Free;
      AfficheLaGrille;
    end;
  end else
  begin
    PGIError('IMPOSSIBLE : Ce sous-traitant est affecté dans le document.',Ecran.caption);
  end;
  //
  //FV1 - 17/01/2018 : FS#2798 - TESTS BL : en affectation ST dans devis, Erreur si clic sur la loupe et pas de ST
  iF TOBSSTRAIT.Detail.count = 0 THEN
  Begin
    BMODIFIER.Visible := False;
    BDelete.Visible   := False;
  end;

end;

function TOF_BTPIECEINTERVMUL.AssignedSSTrait(TOBT: TOB): boolean;
begin
	Result := (TOBPieceTrait.FindFirst(['BPE_FOURNISSEUR'],[TOBT.GetValue('BPI_TIERSFOU')],True) <> nil);
end;

Initialization
  registerclasses ( [ TOF_BTPIECEINTERVMUL ] ) ;
end.
