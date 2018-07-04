{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/04/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMULAFFECTEXTE ()
Mots clefs ... : TOF;BTMULAFFECTEXTE
*****************************************************************}
Unit BTMULAFFECTEXTE_TOF ;

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
     HTB97,
     HMsgBox,
     vierge,
     menus,
     UentCommun,
     UTOF,
     Graphics,
     Grids,
     UCotraitance;

Type
	TModeTrait = (TmtAffaire,TmtDocument);
  TTypeProvenance = (ttpDevis,TtpFacture,TtpAffaire,ttpUnknonw);

  TOF_BTMULAFFECTEXTE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	FFromMem    : Boolean;
    fprovenance : TTypeProvenance;
    fselection  : Boolean;
    Ok_SSTraite : Boolean;
    //
  	fModeTrait  : TmodeTrait;
    Fgestion    : Tmodegestion;
  	fAffaire    : string;
    Fcledoc     : R_CLEDOC;
  	LATOB,TOBL  : TOB;
    TOBPiecesTraitPrec : TOB;
  	GS          : THGrid;
    BEXPORT     : TToolbarButton97;
    //
    fColNames   : string;
    Ftitre      : Hstring;
    FLargeur    : string;
    FAlignement : string;
    ColPa :   Integer;
		MnExportDocXls : TmenuItem;
    //
  	procedure getComponents;
  	procedure decodeArgs (S : string);
    procedure ChargeLaGrille;
    procedure ChargelaTOB;
    procedure Definigrid;
    procedure DefiniGrille(grid: ThGrid; fColNames, Falignement, Ftitre,fLargeur: string);
    procedure MnExportDocXlsClick (Sender : Tobject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSDblClick (Sender : Tobject);
    procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);


  end ;

Implementation
uses UfactExportXLS,
      windows,
      Messages;

procedure TOF_BTMULAFFECTEXTE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFECTEXTE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFECTEXTE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFECTEXTE.OnLoad ;

begin
  Inherited ;
end ;

procedure TOF_BTMULAFFECTEXTE.OnArgument (S : String ) ;
var cancel : boolean;
begin
  Inherited ;
	TOBPiecesTraitPrec := TOB.Create ('LES INFOS REGL',nil,-1);
  FFromMem := false;
  fselection := false;
  getComponents;
  decodeArgs (S);
  ColPa := -1;
  if Fgestion =TmgSousTraitance then ecran.caption := 'Liste des Sous-traitants'
  else if Fgestion =TmgCotraitance then ecran.caption := 'Liste des Cotraitants'
  else Ecran.Caption := 'Liste des Intervenants';

  if fSelection then GS.OnDblClick := GSDblClick;

  ChargelaTOB;

  //FV1 : si aucun cotraitant on n'affiche pas le récap aucun intérêt
  if (laTOB= nil) or (laTOB.detail.count = 0) then
  begin
  	PGIInfo ('Aucune information actuellement', Ecran.Caption);
    PostMessage(TFVierge(Ecran).Handle, WM_CLOSE, 0, 0);
  end
  else
  begin
    Definigrid;
    ChargeLaGrille;
    GSRowEnter(self,1,cancel,false);
  end;

end ;


procedure TOF_BTMULAFFECTEXTE.ChargelaTOB;
VAR Indice : integer;
		MontantTot,MtFacture,MtRegle : double;
begin

  if TFvierge(ecran).LaTOF.LaTOB <> nil then
  begin
  	LATOB := TFvierge(ecran).LaTOF.LaTOB; // passé via mémoire (depuis document)
    if laTOB.detail.count = 0 then exit;
    FFromMem := true;
    if (pos(LaTOB.detail[0].GetString('BPE_NATUREPIECEG'),'DBT;ETU')>0) then fprovenance := ttpDevis
    else if (Pos(LaTOB.detail[0].GetString('BPE_NATUREPIECEG'),'FBT;FBP')>0) then fprovenance := TtpFacture
    else fprovenance  := ttpUnknonw;
  end else
  begin
  	laTOB := TOB.Create('LES DONNEES',nil,-1);
    
		if fModeTrait = TmtAffaire  then fprovenance:= TtpAffaire
    else if pos(Fcledoc.NaturePiece,'DBT;ETU;')> 0  then fprovenance := ttpDevis
    else if (Pos(Fcledoc.NaturePiece ,'FBT;FBP')>0) then fprovenance := TtpFacture
    else fprovenance := ttpUnknonw;

    if fModeTrait = TmtDocument then
    begin
      if Fgestion = TmgCoTraitance then
      	LoadLaTOBPieceTrait (laTOB,Fcledoc, 'COTRAITE')
      else if Fgestion =TmgSousTraitance then
       	LoadLaTOBPieceTrait (laTOB,Fcledoc, 'SSTRAITE')
      else
      	LoadLaTOBPieceTrait (laTOB,Fcledoc, '');
    end else
    begin
      ConstitueCumuleAnterieurAffaire (FAffaire ,TOBPiecesTraitPrec);
      if Fgestion = TmgCoTraitance then
       	LoadLaTOBPieceAffaireTrait (laTOB,Faffaire, 'COTRAITE')
      else if Fgestion =TmgSousTraitance then
       	LoadLaTOBPieceAffaireTrait (laTOB,Faffaire, 'SSTRAITE')
      else
      	LoadLaTOBPieceAffaireTrait (laTOB,Faffaire, '');
    end;
  end;

  if (laTOB= nil) or (laTOB.detail.count = 0) then exit;

  MontantTot := 0;
  For Indice := 0 to LaTOB.detail.count -1 do
  begin
  	MontantTot := MontantTot+LATOB.detail[Indice].GetDouble('BPE_TOTALHTDEV');
  end;

  Ok_SSTraite := False;

  For Indice := 0 to LaTOB.detail.count -1 do
  begin
    if Fgestion = TmgIntervenant then //contrôle si au moins une ligne sous-traitance
    begin
      if Latob.detail[Indice].GetString('BPE_TYPEINTERV') = 'Y00' then Ok_SSTraite := True
    end;

    if not LATOB.detail[0].FieldExists('SEL') then
    begin
    	LATOB.detail[Indice].AddChampSupValeur ('SEL','');
    end;

    if not LATOB.detail[Indice].FieldExists('LIBELLE') then
    begin
      LATOB.detail[Indice].AddChampSupValeur ('LIBELLE',rechdom('GCTIERSFOURN',
                                                                 laTOB.detail[Indice].getvalue('BPE_FOURNISSEUR'),
                                                                 false));
    end;

    if not LATOB.detail[Indice].FieldExists('POURCENT') then
    begin
      LATOB.detail[Indice].AddChampSupValeur ('POURCENT',0);
    end;

    if MontantTot <> 0 then
    begin
      LATOB.detail[Indice].PutValue ('POURCENT',arrondi((LATOB.detail[Indice].GetValue ('BPE_TOTALHTDEV')/MontantTot)*100,2));
    end;

    if fprovenance = TtpFacture then
    begin
			LATOB.detail[Indice].PutValue ('REGLE',LATOB.detail[Indice].GetValue ('BPE_MONTANTREGL'));
    end else if fprovenance = TtpAffaire then
    begin
      if Latob.detail[Indice].GetString('BPE_TYPEINTERV')='Y00' then
      begin

      end else
      begin
        GetReglementsPrec (Latob.detail[Indice].GetString('BPE_FOURNISSEUR'),TOBPiecesTraitPrec,false,MtFacture,MtRegle);
				LATOB.detail[Indice].PutValue ('FACTURE',MtFacture);
				LATOB.detail[Indice].PutValue ('REGLE',MtRegle);
      end;
    end;
  end;
  
end;

procedure TOF_BTMULAFFECTEXTE.ChargeLaGrille;
begin
	GS.RowCount := LATOB.Detail.count+1;
  if GS.rowCount < 2 then GS.rowCount := 2;
	LATOB.PutGridDetail(GS,false,false,fColNames);
  TFvierge(ecran).HMTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTMULAFFECTEXTE.OnClose ;
begin
  Inherited ;
  if not FFromMem then laTOB.free;
  TOBPiecesTraitPrec.free;
end ;

procedure TOF_BTMULAFFECTEXTE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFECTEXTE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFECTEXTE.decodeArgs(S: string);
var lesArgs,UnArg,UnChamp,UneValeur : string;
begin
  LesArgs := S;

  Fgestion := Tmgintervenant;

  Repeat
    UnArg := READTOKENST(lesArgs);
    UnChamp := '';
    Unevaleur := '';
    if UnArg <> '' then
    begin
    	if Pos('=',UnArg) > 0 then
      begin
      	UnChamp := copy(UnArg,1,Pos('=',UnArg)-1);
        Unevaleur := copy(UnArg,Pos('=',UnArg)+1,255);
      end else
      begin
      	unChamp := UnArg;
      end;
      if Unchamp = 'COTRAITANCE' then
      begin
      	Fgestion := TmgCotraitance;
      end else if UnChamp = 'SOUSTRAITANCE' then
      begin
      	Fgestion := TmgSousTraitance;
      end else if UnChamp = 'AFFAIRE' then
      begin
      	fAffaire := Unevaleur;
        fModeTrait := TmtAffaire;
      end else if UnChamp = 'DOCUMENT' then
      begin
      	DecodeCleDoc (Unevaleur,Fcledoc);
        fModeTrait := TmtDocument;
      end else if UnChamp = 'SELECTION' then
      begin
        fselection := true;
      end;
    end;
  Until UnArg = '';
  if fAffaire <> '' then BEXPORT.visible := true else BEXPORT.visible := false;
end;

procedure TOF_BTMULAFFECTEXTE.getComponents;
begin
	GS := THGrid(GetControl('GS'));
  GS.OnRowEnter := GSRowEnter;
  GS.PostDrawCell := GSPostDrawCell;
  BEXPORT := TToolbarButton97 (getControl('BEXPORT'));
  MnExportDocXls := TmenuItem(getControl('MnExportDocXls'));
  MnExportDocXls.OnClick := MnExportDocXlsClick;
end;

procedure TOF_BTMULAFFECTEXTE.DefiniGrille (Grid: ThGrid; fColNames,Falignement,Ftitre,fLargeur : string);
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,lelement : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec,NbCols : integer;
    indice,ind : integer;
    FFQTE,FFDEC : string;
begin
  //
  if Grid.rowcount < 2 then Grid.rowCount := 2;
  st := fColNames;
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
  st := fColNames;
  lesalignements := Falignement;
  lestitres := Ftitre;
  leslargeurs := fLargeur;

  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    OkVisu  := True;
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    grid.cells[Indice,0] := leTitre;
    grid.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G'      then grid.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then grid.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then grid.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if (Nam = 'BPE_MONTANTPA') then ColPa := indice;
    if Fgestion = TmgIntervenant then //si pas de sous traitance on affiche pas le montant d'achat
      if (Nam = 'BPE_MONTANTPA') and (not Ok_SSTraite) then OkVisu := Ok_SSTraite;
    //
    if fprovenance = TtpDevis then
    begin
			if pos(Nam ,'FACTURE;REGLE;') > 0 then OkVisu := false;
    end else if fprovenance = TtpFacture then
    begin
			if pos(Nam ,'FACTURE') > 0 then OkVisu := false;
    end;
    //
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
      if Sep then FFQte := FFQTE+' ##';
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

procedure TOF_BTMULAFFECTEXTE.Definigrid;
begin
	if Fgestion = TmgCotraitance then
  begin
    // Définition de la liste de saisie pour la grille Détail
    fColNames := 'SEL;BPE_FOURNISSEUR;LIBELLE;BPE_TOTALHTDEV;BPE_TOTALTTCDEV;POURCENT;FACTURE;REGLE;';
    Falignement := 'D.0  ---;G.0  ---;G.0  ---;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;';
    Ftitre := 'Num;Cotraitant;Raison sociale;Montant HT;Montant TTC;%;Facturé;En Règlement;';
    fLargeur := '5;18;60;20;20;12;20;20;';
  end
  else if Fgestion = TmgSousTraitance then
  begin
    fColNames := 'SEL;BPE_FOURNISSEUR;LIBELLE;BPE_TOTALHTDEV;BPE_MONTANTPA;POURCENT;FACTURE;REGLE;';
    Falignement := 'D.0  ---;G.0  ---;G.0  ---;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;';
    Ftitre := 'Num;Sous-Traitant;Raison sociale;Montant HT;Mt Achat;%;Facturé;En Règlement;';
    fLargeur := '5;18;60;20;20;12;20;20;';
  end
  else
  begin
    fColNames   := 'SEL;BPE_TYPEINTERV;BPE_FOURNISSEUR;LIBELLE;BPE_MONTANTPA;BPE_TOTALHTDEV;BPE_TOTALTTCDEV;POURCENT;FACTURE;REGLE;';
    Falignement := 'D.0  ---;C.0O --X;G.0  ---;G.0  ---;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;';
    Ftitre := 'Num;Type;Intervenant;Raison sociale;Ss-traitance TTC;Montant HT;Montant TTC;%;Facturé;En Règlement;';
    fLargeur := '5;5;18;60;20;20;20;12;20;20;';
  end;
  //
  DefiniGrille (GS,fColNames,Falignement,Ftitre,fLargeur);
  //
end;


procedure TOF_BTMULAFFECTEXTE.MnExportDocXlsClick(Sender: Tobject);
var Fournisseur : string;
    TOBL : TOB;
    cledoc : r_cledoc;
begin
  fillchar(cledoc,sizeof(cledoc),0);
  TOBL := LATOB.Detail[GS.Row-1];
  Fournisseur := TOBL.getvalue('BPE_FOURNISSEUR');
  LanceExportDocument (fAffaire,Fournisseur);
end;

procedure TOF_BTMULAFFECTEXTE.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  BEXPORT.enabled := false;
 	if Ou <= GS.fixedRows then exit;
  if laTOB.detail.count = 0 then exit;
  TOBL := LaTOB.Detail[Ou-1]; if TOBl = nil then exit;
  if TOBL.getvalue('BPE_FOURNISSEUR')<> '' then Bexport.Enabled := true;
end;

procedure TOF_BTMULAFFECTEXTE.GSDblClick(Sender: Tobject);
var Fournisseur : string;
    TOBL : TOB;
begin
  TOBL := LATOB.Detail[GS.Row-1];
  Fournisseur := TOBL.getvalue('BPE_FOURNISSEUR');
  TFVierge(ecran).Retour := Fournisseur;
  TFvierge(ecran).Close;
end;

procedure TOF_BTMULAFFECTEXTE.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var Arect : TRect;
begin
	if (Acol < GS.FixedCols) or (Arow < gs.FixedRows) then exit;
  ARect := GS.CellRect(ACol, ARow);
  //
  if laTOB.detail.count = 0 then exit;
  TOBL := LaTOB.Detail[ARow-1]; if TOBl = nil then exit;
  if (ColPa > 0) and (ACol = ColPa) then
  begin
    if TOBL.GetValue('BPE_TYPEINTERV') <> 'Y00' then
    begin
    	Canvas.FillRect(ARect);
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTMULAFFECTEXTE ] ) ;
end.
