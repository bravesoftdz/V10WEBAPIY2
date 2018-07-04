{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/06/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISIEPAGEMETRE ()
Mots clefs ... : TOF;BTSAISIEPAGEMETRE
*****************************************************************}
Unit BTSAISIEPAGEMETRE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     vierge,
{$IFNDEF EAGLCLIENT}
     db,
     fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     mainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AglInit,
     UTOF,
     HTB97,
     windows,
     messages,
     UentCommun ;

Type
  TOF_BTSAISIEPAGEMETRE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	Validated : boolean;
  	Action : TActionFiche;
  	TOBmetres,TOBmetres_O : TOB;
    GS : THGrid;
    stprev : string;
    Total : double;
    RESULTAT : THNumEdit;
    Iindice,Ilibelle,Ioperation,Ivaleur : integer;
    fListe,fColonnes : string;
    FNomTable, FLien, FSortBy, stCols : string;
    Ftitre : Hstring;
    FLargeur, FAlignement, FParams, FPerso: string;
    title, NC : Hstring;
    OkTri, OkNumCol : boolean;
    nbcolsinliste : integer;
    //
  	procedure Affichelaligne (Indice : integer);
    procedure GetComponents;
    procedure SetparamGrille;
    function getTOBLigneMetre (Ligne : integer) : TOB;
    procedure AfficheTout;
    procedure DefiniGrille;
    procedure SetEvents (Etat: boolean);
    procedure SetLigne (Arow:integer;Acol : integer=-1);
    // gestion grille de saisie
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    //
    procedure AffichelaGrille(from : integer =1);
    procedure ZoneSuivanteOuOk(Grille: THgrid; var ACol, ARow: integer;
      var Cancel: boolean);
    function ZoneAccessible(Grille: THgrid; var ACol,
      ARow: integer): boolean;
    procedure CalculeTotalisation(PrevValue: double; PrevOperateur: string;
      NewValue: double; NewOperateur: string);
    procedure SetCalculLigne(TOBL: TOB; Colonne: integer;
      ContenuCellule: string);
    procedure AddLigne;
    procedure CalculeSousTotal(TOBL: TOB; ligne: integer);
    procedure NumeroteLignes;
    procedure getTotalisation;
    procedure SetTotalisation;
    procedure ClickInsert(ARow : integer);
    procedure SupLigneClick(Arow : integer);
    procedure recalculSousTotaux;

    // -----
  end ;

function AddMetresLigne (TOBL,TOBMetres : TOB; Prefixe : String) : TOB; overload;
function AddMetresLigne (TOBMetres : TOB; NumOrdre, UniqueBlo : integer; Prefixe : String) : TOB; overload;
procedure loadlesMetres (cledoc : r_cledoc; TOBmetres : TOB);
function GereMetre (TOBL,TOBmetres : TOB; Action : TactionFiche; Prefixe : String) : double;
procedure ValidelesMetres (TOBPiece,TOBmetres,TobOuvrage : TOB);
function IsMetreExist(TOBmetres: TOB; numOrdre : integer) : boolean;

Implementation
uses UCalculsMetre,UtilTOBPiece, Paramsoc;

function IsMetreExist(TOBmetres : TOB; numOrdre : integer) : boolean;
begin
	result := (TOBmetres.FindFirst(['NUMORDRE'],[NumOrdre],true) <> nil);
end;

function AddMetresLigne (TOBMetres : TOB; NumOrdre, UniqueBlo : integer; Prefixe : string) : TOB;
begin
  result := TOB.Create ('UN METRE',TOBmetres,-1);

  result.AddChampSupValeur('PREFIXE',   Prefixe);
  result.AddChampSupValeur('NUMORDRE',  NumOrdre);
  result.AddChampSupValeur('UNIQUEBLO', UniqueBlo);
  result.AddChampSupValeur('RESULTAT',  0);
  result.AddChampSupValeur('USED',      'X');
end;

function AddMetresLigne (TOBL,TOBMetres : TOB; Prefixe : String) : TOB;
var numOrdre  : integer;
    UniqueBlo : Integer;
begin

  NumOrdre  := 0;
  UniqueBlo := 0;

  if prefixe = '' then Prefixe := 'GL';

	if TOBL <> nil then
  begin
    NumOrdre  := StrtoInt(TOBL.GetValue(Prefixe + '_NUMORDRE'));
    if Prefixe = 'BLO' then UniqueBlo := StrtoInt(TOBL.GetValue('BLO_UniqueBLO'));
  end;

  result := AddMetresLigne(TOBMetres,NumOrdre, UniqueBlo, Prefixe);

end;

procedure loadlesMetres (cledoc: r_cledoc; TOBmetres : TOB);
var TOBInter      : TOB;
		TOBII         : TOB;
    TOBPAPA       : TOB;
    LastNumOrdre  : integer;
    LastUniqueBlo : Integer;
    Q             : Tquery;
    Indice        : integer;
begin

	TOBPAPA := nil;

	LastNumOrdre := 0;
  //LastUniqueBlo:= 0;

	TOBInter := TOB.Create('LES METRES PLAT',nil,-1);

  TRY
    Q := OpenSQL('SELECT * FROM BLIGNEMETRE WHERE ' + WherePiece(CleDoc, ttdLigneMetre, False), True,-1, '', True);
    TOBInter.LoadDetailDB('BLIGNEMETRE', '', '', Q, False);
    Ferme(Q);
    //
    if TOBinter.detail.count > 0 then
    begin
      indice := 0;
      repeat
        TOBII := TOBInter.detail[Indice];
        if TOBII.GetValue('BLM_NUMORDRE')<>LastNumordre then
        begin
          LastNumOrdre  := TOBII.GetValue('BLM_NUMORDRE');
          LastUniqueBlo := TOBII.GetValue('BLM_UNIQUEBLO');
          if LastUniqueBlo = 0 then
            TOBPAPA := AddMetresLigne(TOBmetres,LastNumOrdre,LastUniqueBlo,'GL')
          else
            TOBPAPA := AddMetresLigne(TOBmetres,LastNumOrdre,LastUniqueBlo,'BLO')
        end;
        if TOBPAPA <> nil then
        begin
          TOBII.ChangeParent(TOBPAPA,-1);
          if TOBII.getValue('BLM_OPERATION')='+' then
          begin
            TOBPAPA.putValue('RESULTAT',TOBPAPA.getValue('RESULTAT')+TOBII.getValue('BLM_VALEUR'))
          end
          else if TOBII.getValue('BLM_OPERATION')='-' then
          begin
            TOBPAPA.putValue('RESULTAT',TOBPAPA.getValue('RESULTAT')-TOBII.getValue('BLM_VALEUR'))
          end;
        end;
      until Indice >= TOBInter.detail.count;
    end;
  FINALLY
  	FreeAndNil(TOBinter);
  END;
  //
end;

function GereMetre (TOBL,TOBMetres : TOB; Action : tActionFiche; Prefixe : String) : double;
var TOBM : TOB;
begin
	result := 0;

  if TOBMetres = nil then
    TOBM := AddmetresLigne (TOBL,TOBMetres,Prefixe)
  else
  Begin
    if Prefixe = 'GL' then
  	  TOBM := TOBmetres.FindFirst(['NUMORDRE'],[TOBL.GetValue('GL_NUMORDRE')],true)
    else
      TOBM := TOBmetres.FindFirst(['NUMORDRE', 'UNIQUEBLO'],[TOBL.GetValue('BLO_NUMORDRE'), TOBL.GetValue('BLO_UNIQUEBLO')],true);
    //
    if TOBM = Nil then TOBM := AddmetresLigne (TOBL,TOBMetres,Prefixe);
  end;

	TheTOB := TOBM;
  AGLLanceFiche('BTP','BTSAISIEPAGEMETRE','','','ACTION='+ActionToString(Action) );

  TheTOB := nil;
  if TOBM.detail.count = 0 then
  begin
  	TOBM.Free;
  end
  else
  begin
  	if action <> taconsult then result := TOBM.GetValue('RESULTAT');
  end;

  if not GetParamSocSecur('SO_QTEAZERO', False) then
  begin
    if Result = 0 then Result := 1;
  end;

end;

procedure ValidelesMetres (TOBPiece,TOBmetres,TobOuvrage : TOB);

	procedure SetinfosMetres (TOBpiece,TOBM : TOB);
  var indice : integer;
  		TOBDM : TOB;
  begin
		for indice := 0 to TOBM.detail.count -1 do
    begin
    	TOBDM := TOBM.Detail[indice];
      TOBDM.PutValue('BLM_NATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'));
      TOBDM.PutValue('BLM_SOUCHE',      TOBPiece.getValue('GP_SOUCHE'));
      TOBDM.PutValue('BLM_NUMERO',      TOBPiece.getValue('GP_NUMERO'));
      TOBDM.PutValue('BLM_INDICEG',     TOBPiece.getValue('GP_INDICEG'));
      TOBDM.PutValue('BLM_NUMORDRE',    TOBM.getValue('NUMORDRE'));
      TOBDM.PutValue('BLM_UNIQUEBLO',   TOBM.getValue('UNIQUEBLO'));
    end;
  end;

var indice : integer;
		TOBM : TOB;
begin

	for indice := 0 to TOBmetres.detail.count -1 do
  begin
    TOBM := TOBmetres.detail[indice];
    if TOBM.GetValue('UNIQUEBLO')=0 then
    begin
      if TOBPiece.FindFirst(['GL_NUMORDRE'],[TOBM.getValue('NUMORDRE')],true) = nil then
        TOBM.PutValue('USED','-')
      else
       SetinfosMetres (TOBpiece,TOBM);
    end
    else
    begin
      if TOBOuvrage.FindFirst(['BLO_NUMORDRE', 'BLO_UNIQUEBLO'],[TOBM.getValue('NUMORDRE'),TOBM.getValue('UNIQUEBLO')],true) = nil then
        TOBM.PutValue('USED','-')
      else
        SetinfosMetres (TOBpiece,TOBM);
    end;
  end;

  // seconde phase (nettoyage)
  if TOBmetres.detail.count > 0 then
  begin
    Indice := 0;
    repeat
      if TOBmetres.detail[Indice].getValue('USED')='-' then TOBmetres.detail[Indice].free else Inc(Indice);
    until Indice >= TOBmetres.detail.count;
  end;

end;

{ TOF_BTSAISIEPAGEMETRE }

procedure TOF_BTSAISIEPAGEMETRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEPAGEMETRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEPAGEMETRE.OnUpdate ;
begin
  Inherited ;
	Validated := true;  
end ;

procedure TOF_BTSAISIEPAGEMETRE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEPAGEMETRE.OnArgument (S : String ) ;
begin
  Inherited ;
	Validated := false;
  TOBmetres_O := AddMetresLigne (nil,nil,'');
	Total := 0;
  Action := taModif;
  fListe := 'BLIGNEMETRE';
  //
  GetComponents;
  TOBmetres   := LaTOB;
  TOBmetres_O := AddMetresLigne (nil,nil,'');
  TOBmetres_O.dupliquer(TOBmetres,true,true);
  if TOBmetres.detail.count = 0 then AddLigne;
  NumeroteLignes ;
	getTotalisation;
  SetparamGrille;
	AfficheTout;
  SetEvents(true);
  SetLigne (1);
end ;

procedure TOF_BTSAISIEPAGEMETRE.OnClose ;
begin
  Inherited ;
  if not Validated then
  begin
  	TOBMetres.ClearDetail;
		TOBMetres.Dupliquer(TOBmetres_O,true,true);
  end;
  TOBmetres_O.Free;
end ;

procedure TOF_BTSAISIEPAGEMETRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEPAGEMETRE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEPAGEMETRE.GetComponents;
begin
	GS := THgrid(GetControl('GS'));
  RESULTAT := THNumEdit(GetControl('RESULTAT'));
  RESULTAT.Decimals := V_PGI.OkDecQ;
end;

procedure TOF_BTSAISIEPAGEMETRE.AffichelaGrille (from : integer =1);
var Indice : integer;
begin
	if from < 1 then from := 1;
  GS.RowCount := TOBmetres.detail.count +2;
  for Indice := from-1 to TOBmetres.detail.count -1 do
  begin
  	Affichelaligne (Indice+1);
  end;
end;

procedure TOF_BTSAISIEPAGEMETRE.Affichelaligne(Indice: integer);
var TOBL : TOB;
begin
	TOBL := getTOBLigneMetre (Indice);
  TOBL.PutLigneGrid(GS,Indice,false,false,fColonnes)
end;


procedure TOF_BTSAISIEPAGEMETRE.SetparamGrille;
var lelement,laliste : string;
begin
  // récupération du paramétrage général des grilles

  ChargeHListe (fListe, FNomTable, FLien, FSortBy, stCols, FTitre,
                FLargeur, FAlignement, FParams, title, NC, FPerso, OkTri, OkNumCol);
  fColonnes  := stCols;
  laliste := stcols;
  nbcolsInListe := 0;
  repeat
    lelement := READTOKENST (laliste);
    if lelement <> '' then
    begin
      if lelement = 'BLM_INDICE' then
      begin
      	IIndice := nbcolsinListe
      end else if lelement = 'BLM_CHAINE' then
      begin
      	Ilibelle := nbcolsinListe
      end else if lelement = 'BLM_OPERATION' then
      begin
      	IOperation := nbcolsinListe
      end else if lelement = 'BLM_VALEUR' then
      begin
      	Ivaleur := nbcolsinListe;
      end;
      inc(nbcolsInListe);
    end;
  until lelement = '';
end;

procedure TOF_BTSAISIEPAGEMETRE.SetEvents(Etat: boolean);
begin
	if Etat then
  begin
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
    GS.OnKeyDown := GSKeyDown;
  end else
  begin
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnCellEnter := nil;
    GS.OnCellExit := nil;
    GS.OnKeyDown := nil;
  end;
end;

procedure TOF_BTSAISIEPAGEMETRE.AfficheTout;
begin
  GS.VidePile(false);
  DefiniGrille;
  if TOBmetres.detail.count = 0 then GS.rowCount := 2 else GS.rowCount := TOBmetres.detail.count+2;
  AffichelaGrille;
  RESULTAT.Value := TOBmetres.getValue('RESULTAT');
  //
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
end;

procedure TOF_BTSAISIEPAGEMETRE.DefiniGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
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

  st := fColonnes;
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
    GS.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if OkVisu then
    begin
    	GS.ColWidths[indice] := strtoint(lalargeur);
    end else
    begin
    	GS.ColWidths[indice] := -1;
    end;

    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then GS.ColDrawingModes[Indice]:= 'IMAGE';
    end
    else if (Dec<>0) or (Sep) then
    begin
    	if OkNulle then
      begin
				FFLQte := FFQTE+';-'+FFQTE+'; ;';
      	GS.ColFormats[indice] := FFLQTE ;
      end else
      begin
      	GS.ColFormats[indice] := FF ;
      end;
    end;
  end ;
	// fige les colonnes
  GS.ColEditables [Ivaleur] := false;
  GS.ColLengths  [Ioperation] := 1;
  //
end;


function TOF_BTSAISIEPAGEMETRE.getTOBLigneMetre(Ligne: integer) : TOB;
begin
  result := nil;
  if Ligne > TOBmetres.detail.count then exit;
	result := TOBmetres.detail[Ligne-1];
end;

procedure TOF_BTSAISIEPAGEMETRE.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  if not Cancel then
  begin
    stprev := GS.Cells [Acol,Arow];
  end;
end;

procedure TOF_BTSAISIEPAGEMETRE.getTotalisation ;
begin
	RESULTAT.Value := TOBmetres.getValue('RESULTAT');
  ToTal := TOBmetres.getValue('RESULTAT');
end;

procedure TOF_BTSAISIEPAGEMETRE.SetTotalisation;
begin
	RESULTAT.Value := Total;
  TOBmetres.PutValue('RESULTAT',Total);
end;

procedure TOF_BTSAISIEPAGEMETRE.GSCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  if Action = taConsult then Exit;
  if (stPrev <> GS.cells[Acol,Arow]) or (GS.cells[Acol,Arow]= '') then
  begin
    TOBL := getTOBLigneMetre (Arow);
    if TOBL = nil then exit;
    //
    if ACol = Ilibelle  then  // Saisie de la formule ou du libelle
    Begin
    	SetCalculLigne (TOBL,Acol,GS.cells[Acol,Arow]);
      SetTotalisation;
    end else
    if ACol = Ioperation  then  // opération à effectuer sur la ligne (+,-)
    Begin
    	if pos (GS.cells[Acol,Arow],'+;-;=')= 0 then
      begin
        PgiError ('Les opérateurs autorisés sont : + - ou =');
        cancel := true;
        exit;
      end;
      if (GS.cells[Acol,Arow] = '=') and (TOBL.getValue('BLM_ISFORMULE')='X') then
      begin
        PgiError ('Seul les opérateurs + et - sont autorisés sur une formule de calcul');
        cancel := true;
        exit;
      end;
      if Pos (GS.cells[Acol,Arow],'+;-')> 0 then
      begin
        SetCalculLigne (TOBL,Acol,GS.cells[Acol,Arow]);
      	SetTotalisation;
      end else
      begin
      	TOBL.PutValue('BLM_OPERATION','=');
				CalculeSousTotal (TOBL,Arow);
      end;
    end;
    if not cancel then
    begin
      stPrev := GS.cells[Acol,Arow];
			Affichelaligne (Arow);
    end;
  end;
  if not cancel then
  begin
  	recalculSousTotaux;
  end;
end;

procedure  TOF_BTSAISIEPAGEMETRE.SetCalculLigne (TOBL : TOB; Colonne : integer; ContenuCellule: string);
var PrevValue,newvalue : double;
		PrevOperateur,newOperateur : string;
begin
	if TOBL = nil then exit;
  PrevValue := TOBL.GetValue('BLM_VALEUR');
  Prevoperateur := TOBL.GetValue('BLM_OPERATION');
	if Colonne = Ilibelle then
  begin
    TOBL.PutValue('BLM_CHAINE',ContenuCellule);
  	if PrevOperateur = '=' then exit;
    newvalue := Arrondi(CalculeFormuleLigne(ContenuCellule),V_PGI.okdecQ);
    if NewValue <> 0 then TOBL.PutValue('BLM_ISFORMULE','X') else TOBL.PutValue('BLM_ISFORMULE','-');
    TOBL.PutValue('BLM_VALEUR',NewValue);
    newoperateur := TOBL.GetValue('BLM_OPERATION');
  	CalculeTotalisation (PrevValue,PrevOperateur,NewValue,NewOperateur);
  end else if Colonne = Ioperation then
  begin
  	newValue := TOBL.GetValue('BLM_VALEUR');
    TOBL.PutValue('BLM_OPERATION',ContenuCellule);
    newOperateur := ContenuCellule;
  	CalculeTotalisation (PrevValue,PrevOperateur,NewValue,NewOperateur);
  end;
end;

procedure TOF_BTSAISIEPAGEMETRE.CalculeTotalisation (PrevValue : double;PrevOperateur : string;NewValue : double; NewOperateur : string);
begin
	if prevOperateur = '+' then
  begin
  	// Dans ce cas on déduit l'ancienne valeur pour ajouter la nouvelle valeur
    Total := Total - PrevValue;
  end else if prevOperateur = '-' then
  begin
  	// Dans ce cas on ajoute l'ancienne valeur pour déduire la nouvelle valeur
    Total := Total + PrevValue;
  end;
  //
  if NewOperateur = '+' then
  begin
    Total := Total + newValue;
  end else if newOperateur = '-' then
  begin
    Total := Total - newValue;
  end;
end;


procedure TOF_BTSAISIEPAGEMETRE.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
  if Action = taConsult then Exit;
  if Ou <= GS.fixedCols then exit;
  if Ou > TOBmetres.detail.count then
  begin
  	AddLigne;
  	NumeroteLignes ;
    GS.rowCount := TOBmetres.detail.count +2;
    Affichelaligne(ou);
  end;
  TOBL := getTOBLigneMetre (Ou);
  if TOBL = nil then exit;
end;

procedure TOF_BTSAISIEPAGEMETRE.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if Action = taConsult then Exit;
end;

procedure TOF_BTSAISIEPAGEMETRE.ZoneSuivanteOuOk(Grille : THgrid;var ACol, ARow : integer;var  Cancel : boolean);
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
        Sens := -1;
        continue;
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

procedure TOF_BTSAISIEPAGEMETRE.SetLigne(Arow, Acol: integer);
var Cancel : boolean;
begin
  GS.synEnabled := false;
  GS.CacheEdit;
  SetEvents (false);
  Arow := Arow;
  Acol := 1;
  GS.Row := Arow;
  GS.Col := Acol;
  GSCellEnter (self,Acol,Arow,cancel);
  stPrev  :=GS.Cells[ACol,ARow] ;
  GS.row := ARow; GS.Col := Acol;
  GS.SynEnabled := true;
  GS.MontreEdit;
  SetEvents(true);
end;

function TOF_BTSAISIEPAGEMETRE.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
begin
  result := false;
  if (Grille.ColWidths[acol] = 0) or (not Grille.ColEditables[acol] ) then exit;
  if Arow < GS.FixedRows then exit;
  if Acol < GS.fixedCols then exit;
  if GS.RowHeights[ARow] <= 0 then Exit;
  result := true;
end;

procedure TOF_BTSAISIEPAGEMETRE.AddLigne;
var TOBD : TOB;
begin
	TOBD := TOB.Create('BLIGNEMETRE',TOBmetres,-1);
  TOBD.putValue('BLM_INDICE',TOBmetres.detail.count);
  TOBD.putValue('BLM_OPERATION','+');
end;

procedure TOF_BTSAISIEPAGEMETRE.CalculeSousTotal (TOBL : TOB;ligne : integer);
var Indice : integer;
		SousTotal : double;
    TOBD : TOB;
begin
	SousTotal := 0;
  if Ligne < 2 then exit;
	for Indice := Ligne -2 downto 0 do
  begin
  	TOBD := TOBmetres.detail[Indice];
    if TOBD.GetValue('BLM_OPERATION') = '=' then break;
    if TOBD.GetValue('BLM_OPERATION') = '+' then
    begin
    	SousTotal := SouStotal + TOBD.GetValue('BLM_VALEUR')
   	end else if TOBD.GetValue('BLM_OPERATION') = '-' then
    begin
    	SousTotal := SouStotal - TOBD.GetValue('BLM_VALEUR')
    end;
  end;
	TOBL.putValue('BLM_VALEUR',SousTotal);
end;

procedure TOF_BTSAISIEPAGEMETRE.NumeroteLignes ;
var indice : integer;
begin
	for Indice := 0 to TOBmetres.detail.count -1 do
  begin
  	TOBMetres.detail[Indice].putvalue('BLM_INDICE',Indice+1);
  end;
end;

procedure TOF_BTSAISIEPAGEMETRE.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var OkG, Vide: Boolean;
begin
  OkG := (Screen.ActiveControl = GS);
  Vide := (Shift = []);
  case Key of
    VK_RETURN: if ((OkG) and (Vide)) then
      begin
        SendMessage(GS.Handle, WM_KEYDOWN , VK_TAB, 0)
      end;
    VK_F10: if Vide then
      begin
        TToolBarButton97(GetControl('Bvalider')).OnClick (self);
      end;
    27:
      begin
        TToolBarButton97(GetControl('Bferme')).OnClick (self);
      end;
    VK_INSERT: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        ClickInsert(GS.Row);
      end;
    VK_DELETE: if ((OkG) and (Shift = [ssCtrl])) then
      begin
        Key := 0;
        SupLigneClick(GS.row);
      end;
  end;
end;

procedure TOF_BTSAISIEPAGEMETRE.ClickInsert(ARow : integer);
var TOBD : TOB;
begin
	TOBD := TOB.Create('BLIGNEMETRE',TOBmetres,Arow-1);
  TOBD.putValue('BLM_OPERATION','+');
	NumeroteLignes;
  GS.CacheEdit;
  GS.SynEnabled := False;
  gs.InsertRow(Arow);
  GS.Row := ARow;
  GS.MontreEdit;
  GS.SynEnabled := True;
	AffichelaGrille (Arow);
end;

procedure TOF_BTSAISIEPAGEMETRE.SupLigneClick(Arow : integer);
var TOBD : TOB;
begin
	TOBD := getTOBLigneMetre(Arow);
  CalculeTotalisation (TOBD.GetValue('BLM_VALEUR'),TOBD.getValue('BLM_OPERATION'),0,'+');
  GS.CacheEdit;
  GS.SynEnabled := False;
  SetEvents(false);
  GS.DeleteRow(ARow);
  TOBD.free;
  if TOBmetres.detail.count = 0 then AddLigne;
  GS.RowCount := TOBmetres.detail.count +2;
  if Arow > TOBmetres.detail.count then Arow := TOBmetres.detail.count;
  GS.MontreEdit;
  GS.SynEnabled := True;
	NumeroteLignes;
	AffichelaGrille (Arow);
  SetTotalisation;
  SetEvents(true);
  SetLigne (Arow);
end;

procedure TOF_BTSAISIEPAGEMETRE.recalculSousTotaux;
var indice : integer;
		TOBL : TOB;
begin
	for indice := 0 to TOBmetres.detail.count -1 do
  begin
  	TOBL := TOBMetres.detail[Indice];
    if TOBL.getvalue('BLM_OPERATION')='=' then
    begin
    	CalculeSousTotal (TOBL,indice+1);
      Affichelaligne(Indice+1);
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTSAISIEPAGEMETRE ] ) ;
end.
