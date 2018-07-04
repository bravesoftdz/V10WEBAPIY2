{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 31/12/2002
Modifié le ... :   /  /
Description .. : CONTROLE FACTURE ACHAT
Suite ........ : Appel en facture fournisseur sur comptabilisation normale
Mots clefs ... : FACTURE;FOURNISSEUR;CONTROLE
*****************************************************************}
Unit GCControleFacture_Tof ;

Interface

Uses StdCtrls, Controls, Classes, AGLInit,
{$IFDEF EAGLCLIENT}
     eMul, MaineAgl,
{$ELSE}
     Mul, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
     EntGC, UTOB, UTOF, HCtrls, HTB97, HEnt1, Graphics, Grids, sysutils,
     hmsgbox, FactUtil, Lookup, ent1, Vierge, FactArticle ;

function Entree_ControleFacture (TobBases, TobBasesSaisies, TobP, TobT : TOB;
                                 Act : TActionFiche;
                                 bSaisieArticle : boolean = False) : boolean;

Type
   TOF_CONTROLEFACTURE = Class (TOF)
   private
      GPIEDBASE : THGrid;
      ColCategorie, ColFamille, ColTaux, ColBase : integer;
      ColValeur, ColBaseSaisie, ColValeurSaisie : integer;
      ColArticle : integer;
      NumLigne : integer;

      bCopier : TToolBarButton97;
      bCopierTTC : TToolBarButton97;
      bNouveau : TToolBarButton97;
      bSupprime : TToolBarButton97;
      bValider : TToolbarButton97;
      //Grid
      procedure EtudieColsListe;
      procedure GPIEDBASE_OnCellEnter(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
      procedure GPIEDBASE_OnCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
      procedure GPIEDBASE_OnElipsisClick (Sender : TObject);
      procedure GPIEDBASE_OnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure GPIEDBASE_OnRowExit(Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
      procedure GPIEDBASE_GetCellCanvas (Acol, ARow : LongInt; Canvas : TCanvas; AState: TGridDrawState);
      procedure TTCSaisie_OnChange (Sender : TObject);
      //Boutons
      procedure CopierClick (Sender : TObject);
      procedure CopierTTCClick (Sender : TObject);
      procedure NouveauClick (Sender : TObject);
      procedure SupprimeClick (Sender : TObject);
      procedure ValiderClick (Sender: TObject);
      //Gestion
      procedure AfficheGrille;
      procedure CalculTTCSaisie;
//      function SelectionCategorieTaxe : string;
      function SelectionFamilleTaxe : string;
      procedure TestCellule;
      procedure ValideLigne(Ou : integer);
      procedure ValideSaisie (stChamp : string; iCol, iRow : integer);
   public
      procedure OnArgument(S : String) ; override ;
      procedure OnClose                ; override ;
  end ;

Implementation

var TobPiedBases, TobPiece, TobTiers : TOB;
    dTotalTTC, dTotalTTCSaisie : double;
    stRegimeTaxe, stRefInterne : string;
    bOk, bSaisie, bArticle : boolean;
    Action : tActionFiche;

function Entree_ControleFacture (TobBases, TobBasesSaisies, TobP, TobT : TOB;
                                 Act : TActionFiche;
                                 bSaisieArticle : boolean = False) : boolean;
var TobPB, TobPBS : TOB;
    iInd : integer;
begin
TobPiedBases := Tob.Create ('', nil, -1);
TobPiedBases.Dupliquer (TobBases, true, true, true);
if TobPiedBases.Detail.Count > 0 then
    begin
    TobPiedBases.Detail[0].AddChampSupValeur ('(BASESAISIE)', 0, true);
    TobPiedBases.Detail[0].AddChampSupValeur ('(VALEURSAISIE)', 0, true);
    end;
for iInd := 0 to TobBasesSaisies.Detail.Count - 1 do
    begin
    TobPBS := TobBasesSaisies.Detail[iInd];
    TobPB := TobPiedBases.FindFirst (['GPB_CATEGORIETAXE', 'GPB_FAMILLETAXE'],
                                     [TobPBS.GetValue ('GPB_CATEGORIETAXE'), TobPBS.GetValue ('GPB_FAMILLETAXE')],
                                     false);
    if TobPB = nil then
        begin
        TobPB := Tob.Create ('PIEDBASE', TobPiedBases, -1);
        TobPB.Dupliquer (TobPBS, true, true, true);
        TobPB.AddChampSupValeur ('(BASESAISIE)', TobPB.GetValue ('GPB_BASEDEV'), false);
        TobPB.PutValue ('GPB_BASEDEV', 0);
        TobPB.AddChampSupValeur ('(VALEURSAISIE)', TobPB.GetValue ('GPB_VALEURDEV'), false);
        TobPB.PutValue ('GPB_VALEURDEV', 0);
        TobPB.AddChampSup ('ENPLUS', false);

        if (bSaisieArticle) and (not TobPB.FieldExists ('(ARTICLE)')) then
            begin
            TobPB.AddChampSupValeur ('(ARTICLE)', '', false);
            TobPB.AddChampSupValeur ('GA_ARTICLE', '', false);
            end;
        end else
        begin
        TobPB.PutValue ('(BASESAISIE)', TobPBS.GetValue ('GPB_BASEDEV'));
        TobPB.PutValue ('(VALEURSAISIE)', TobPBS.GetValue ('GPB_VALEURDEV'));
        end;
    end;

Action := Act;
dTotalTTC := TobP.GetValue ('GP_TOTALTTCDEV');
if not TobBasesSaisies.FieldExists ('TOTALTTC') then
    TobBasesSaisies.AddChampSupValeur ('TOTALTTC', 0, false);
dTotalTTCSaisie := TobBasesSaisies.GetValue ('TOTALTTC');
if not TobBasesSaisies.FieldExists ('REFINTERNE') then
    TobBasesSaisies.AddChampSupValeur ('REFINTERNE', '', false);
//if Action = taConsult then
TobBasesSaisies.PutValue ('REFINTERNE', TobP.GetValue ('GP_REFINTERNE'));
stRefInterne := TobBasesSaisies.GetValue ('REFINTERNE');
stRegimeTaxe := TobP.GetValue ('GP_REGIMETAXE');
bArticle := bSaisieArticle;
TobPiece := TobP;
TobTiers := TobT;

bOk := False;
AglLanceFiche('GC', 'GCCONTROLEFACTURE', '', '', '');

if bOk then
    begin
    TobBasesSaisies.ClearDetail;
    for iInd := 0 to TobPiedBases.Detail.Count - 1 do
        begin
        TobPB := TobPiedBases.Detail[iInd];
        if (TobPB.GetValue ('GPB_CATEGORIETAXE') <> '') and
           (TobPB.GetValue ('GPB_FAMILLETAXE') <> '') then
            begin
            TobPBS := Tob.Create ('PIEDBASE', TobBasesSaisies, -1);
            TobPBS.Dupliquer (TobPB, true, true, true);

            TobPBS.PutValue ('GPB_BASEDEV', TobPB.GetValue ('(BASESAISIE)'));
            TobPBS.PutValue ('GPB_VALEURDEV', TobPB.GetValue ('(VALEURSAISIE)'));
            TobPBS.DelChampSup ('(BASESAISIE)', false);
            TobPBS.DelChampSup ('(VALEURSAISIE)', false);
            end;
        end;
    TobBasesSaisies.PutValue ('TOTALTTC', dTotalTTCSaisie);
    TobBasesSaisies.PutValue ('REFINTERNE', stRefInterne);
    TobP.PutValue ('GP_REFINTERNE', stRefInterne);
    end;

TobPiedBases.Free;
Result := bOk;
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_CONTROLEFACTURE.OnArgument (S : String ) ;
var iInd : integer;
begin
    Inherited ;
    NumLigne := 0;

    bCopier := TToolBarButton97(GetControl('BCopier'));
    bCopier.OnClick := CopierClick;
    bCopierTTC := TToolBarButton97(GetControl('BCopierTTC'));
    bCopierTTC.OnClick := CopierTTCClick;
    bValider := TToolBarButton97(GetControl('BValider'));
    bValider.OnClick := ValiderClick;
    bNouveau := TToolBarButton97(GetControl('BInsert'));
    bNouveau.OnClick := NouveauClick;
    bSupprime := TToolBarButton97(GetControl('BDelete'));
    bSupprime.OnClick := SupprimeClick;

    //Paramètrage du Grid
    GPIEDBASE := THGrid(GetControl('GPIEDBASE'));
    GPIEDBASE.ListeParam := 'GCCONTROLEFACTURE';

    GPIEDBASE.OnRowEnter := GPIEDBASE_OnRowEnter;
    GPIEDBASE.OnRowExit := GPIEDBASE_OnRowExit;
    GPIEDBASE.OnCellEnter := GPIEDBASE_OnCellEnter;
    GPIEDBASE.OnCellExit := GPIEDBASE_OnCellExit;
    GPIEDBASE.OnElipsisClick := GPIEDBASE_OnElipsisClick;
    GPIEDBASE.GetCellCanvas := GPIEDBASE_GetCellCanvas;

    GPIEDBASE.ColFormats [ColBaseSaisie] := '#.##';
    GPIEDBASE.ColFormats [ColValeurSaisie] := '#.##';

    THNumEdit (GetControl ('TTCSAISIE')).OnChange := TTCSAISIE_OnChange;

    EtudieColsListe;

    //Colonnes non saisissable
    if Action = taConsult then
    begin
        GPIEDBASE.Options := GPIEDBASE.Options - [goEditing] + [goRowSelect];
        bNouveau.Visible := False;
        bCopier.Visible := False;
        bCopierTTC.Visible := False;
        bSupprime.Visible := False;
        SetControlEnabled ('REFINTERNE', False);
        SetControlVisible ('TTCSAISIE', False);
        SetControlVisible ('TTTCSAISIE', False);
        GPIEDBASE.ColLengths [ColArticle] := -1;
        GPIEDBASE.ColWidths [ColArticle] := -1;
        GPIEDBASE.ColLengths [Colbasesaisie] := -1;
        GPIEDBASE.ColWidths [Colbasesaisie] := -1;
        GPIEDBASE.ColLengths [Colvaleursaisie] := -1;
        GPIEDBASE.ColWidths [Colvaleursaisie] := -1;
    end else
    begin
        GPIEDBASE.ColLengths[ColTaux] := -1;
        GPIEDBASE.ColLengths[ColBase] := -1;
        GPIEDBASE.ColLengths[ColValeur] := -1;
        if not bArticle then
        begin
            GPIEDBASE.ColLengths [ColArticle] := -1;
            GPIEDBASE.ColWidths [ColArticle] := -1;
        end;
    end;

        //Formatage des zones
    GPIEDBASE.ColFormats[ColBaseSaisie] := '#########0.00';
    GPIEDBASE.ColTypes[ColBaseSaisie] := 'R';
    GPIEDBASE.ColFormats[ColValeurSaisie] := '#########0.00';
    GPIEDBASE.ColTypes[ColValeurSaisie] := 'R';
    GPIEDBASE.ColFormats [ColCategorie] := 'CB=GCCATEGORIETAXE|CC_CODE="TX1" OR CC_CODE="TX2"|';
    GPIEDBASE.ColFormats[ColFamille] := 'UPPER';

    AfficheGrille;

    SetControlText ('TTCCALCULE', FloatToStr (dTotalTTC));
    SetControlText ('TTCSAISIE', FloatToStr (dTotalTTCSaisie));
    SetControlText ('REFINTERNE', stRefInterne);

    TFVierge(Ecran).HMTrad.ResizeGridColumns(GPIEDBASE);
    GPIEDBASE.Col := ColBaseSaisie;
    bSaisie := False;
    CalculTTCSaisie;
    if (Action <> taConsult) and (dTotalTTCSaisie = 0) then
    begin
        for iInd := 1 to GPIEDBASE.RowCount - 1 do
        begin
            GPIEDBASE.Row := iInd;
            CopierClick (Self);
        end;
        CopierTTCClick (Self);
    end;
end;

procedure TOF_CONTROLEFACTURE.OnClose ;
begin
if Action <> taConsult then
    begin
    if bSaisie then
        begin
        if PgiAsk ('Confirmer l''abandon de la saisie', 'Controle Facture Achat') <> mrYes then
            begin
            LastError := 1;
            end;
        end;
      Inherited ;
    end;
if LastError <> 1 then GPIEDBASE.VidePile(False);
end;

{==============================================================================================}
{==================================== Gestion du Grid =========================================}
{==============================================================================================}
procedure TOF_CONTROLEFACTURE.EtudieColsListe ;
Var NomCol,LesCols : String ;
    NumCol : integer ;
begin
LesCols := GPIEDBASE.Titres[0];
NumCol := 1;
Repeat
    NomCol := Uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        if NomCol='GPB_CATEGORIETAXE' then ColCategorie := NumCol
            else if NomCol='GPB_FAMILLETAXE' then ColFamille := NumCol
            else if NomCol='GPB_TAUXTAXE' then ColTaux := NumCol
            else if NomCol='GPB_BASEDEV' then ColBase := NumCol
            else if NomCol='GPB_VALEURDEV' then ColValeur := NumCol
            else if NomCol='(BASESAISIE)' then ColBaseSaisie := NumCol
            else if NomCol='(VALEURSAISIE)' then ColValeurSaisie := NumCol
            else if NomCol='(ARTICLE)' then ColArticle := NumCol
        end ;
    Inc(NumCol) ;
    Until ((LesCols='') or (NomCol='')) ;
end;

procedure TOF_CONTROLEFACTURE.GPIEDBASE_OnCellEnter(
                Sender : TObject; var ACol, ARow : Integer;
                var Cancel : Boolean);
begin
TestCellule;
end;

procedure TOF_CONTROLEFACTURE.GPIEDBASE_OnCellExit(
                Sender : TObject; var ACol, ARow : Integer;
                var Cancel : Boolean);
var stChamp, stArt : string;
    TobP, TobT, TobArt : TOB;
    stControle, stWhere : string;
    iRow : integer;
    bControle : boolean;
    TSql : TQuery;
    RechArt : T_RechArt;
begin
    TobP := TOB(GPIEDBASE.Objects [0, ARow]);
    if not TobP.FieldExists ('ENPLUS') then
    begin
        GPIEDBASE.CellValues [ColCategorie, ARow] := TobP.GetValue ('GPB_CATEGORIETAXE');
        GPIEDBASE.CellValues [ColFamille, ARow] := TobP.GetValue ('GPB_FAMILLETAXE');
    end else
    if //(TobP.FieldExists ('ENPLUS')) and
       (GPIEDBASE.CellValues [ColCategorie, ARow] <> '') and
       (GPIEDBASE.CellValues [ColFamille, ARow] <> '') then
    begin
        TobT := VH^.LaTOBTVA.FindFirst (['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],
                                        [GPIEDBASE.CellValues [ColCategorie, ARow], stRegimeTaxe,
                                         GPIEDBASE.CellValues [ColFamille, ARow]], False) ;
        if TobT <> nil then
            begin
            GPIEDBASE.CellValues [ColTaux, ARow] := TobT.GetValue ('TV_TAUXACH');
            TobP.PutValue ('GPB_TAUXTAXE', TobT.GetValue ('TV_TAUXACH'));
            end;
    end;
    stChamp := '';
    if ACol = ColBaseSaisie then stChamp := '(BASESAISIE)'
    else if ACol = ColValeurSaisie then stChamp := '(VALEURSAISIE)'
    else if (ACol = ColCategorie) then
        begin
        if (GPIEDBASE.CellValues [ColCategorie, ARow] <> '') then
            begin
            stControle := RechDom ('GCCATEGORIETAXE', GPIEDBASE.CellValues [ColCategorie, ARow], False);
            if stControle = '' then
                begin
                GPIEDBASE.Col := ColCategorie;
                GPIEDBASE.CellValues [ColCategorie, ARow] := TobP.GetValue ('GPB_CATEGORIETAXE');
                end;
            end;
        if TobP.GetValue ('GPB_CATEGORIETAXE') <> GPIEDBASE.CellValues [ColCategorie, ARow] then bSaisie := true;
        TobP.PutValue ('GPB_CATEGORIETAXE', GPIEDBASE.CellValues [ColCategorie, ARow]);
        end else if (ACol = ColFamille) then
        begin
        if (GPIEDBASE.CellValues [ColFamille, ARow] <> '') then
            begin
            bControle := False;
            stControle := RechDom ('GCFAMILLETAXE' + Copy (GPIEDBASE.CellValues [ColCategorie, ARow], 3, 1),
                                   GPIEDBASE.CellValues [ColFamille, ARow], False);
            if stControle = '' then
                begin
                bControle := True;
                end else
                begin
                for iRow := 1 to GPIEDBASE.RowCount - 1 do
                    begin
                    if iRow <> GPIEDBASE.Row then
                        begin
                        if (GPIEDBASE.CellValues [ColCategorie, iRow] = GPIEDBASE.CellValues [ColCategorie, GPIEDBASE.Row]) and
                           (GPIEDBASE.CellValues [ColFamille, iRow] = GPIEDBASE.CellValues [ColFamille, GPIEDBASE.Row]) then
                            begin
                            bControle := True;
                            break;
                            end;
                        end;
                    end;
                end;
            if bControle then
                begin
                GPIEDBASE.Col := ColFamille;
                GPIEDBASE.CellValues [ColFamille, ARow] := TobP.GetValue ('GPB_FAMILLETAXE');
                end;
            end else
            begin
            if (GPIEDBASE.CellValues [ColCategorie, ARow] <> '') and (GPIEDBASE.Col <> ColCategorie) then
                GPIEDBASE.Col := ColFamille; // soit on reste sur famille, soit on revient sur categorie
            end;
        if TobP.GetValue ('GPB_FAMILLETAXE') <> GPIEDBASE.CellValues [ColFamille, ARow] then bSaisie := true;
        TobP.PutValue ('GPB_FAMILLETAXE', GPIEDBASE.CellValues [ColFamille, ARow]);
        end else if (ACol = ColArticle) { and
                    (TobP.GetValue ('(ARTICLE)') <> GPIEDBASE.CellValues [ColArticle, ARow]) } then
        begin
        if not TobP.FieldExists ('(ARTICLE)') then
        begin
            TobP.AddChampSupValeur ('(ARTICLE)', '', False);
            TobP.AddChampSupValeur ('GA_ARTICLE', '', False);
        end;
        if TobP.GetValue ('(ARTICLE)') <> GPIEDBASE.CellValues [ColArticle, ARow] then
            begin
            TobArt := Tob.Create ('ARTICLE', nil, -1);
            RechArt := TrouverArticleSQL (TobPiece.GetValue ('GP_NATUREPIECEG'),
                                          GPIEDBASE.CellValues [ColArticle, ARow],
                                          TobPiece.GetValue ('GP_DOMAINE'), '' ,
                                          TobPiece.GetValue ('GP_DATEPIECE'),
                                          TOBArt, TOBTiers) ;
            if RechArt = traAucun then
                begin
                GPIEDBASE.Col := ACol;
                GPIEDBASE.Row := ARow;
                stWhere := ' AND GA_FAMILLETAXE' + Copy (TobP.GetValue ('GPB_CATEGORIETAXE'), 3, 1) +
                         '="' + TobP.GetValue ('GPB_FAMILLETAXE') + '"';
                GetArticleRecherche (GPIEDBASE, 'Recherche Article', '', '', '', stWhere);
                end else
                begin
                GPIEDBASE.CellValues [ColArticle, ARow] := TobArt.GetValue ('GA_ARTICLE');
                end ;
            Ferme(TSql) ;
            TobArt.Free;

            TobP.PutValue ('GA_ARTICLE', GPIEDBASE.CellValues [ColArticle, ARow]);
            end;
        stArt := Copy (GPIEDBASE.CellValues [ColArticle, ARow], 1, 17);
        TobP.PutValue ('(ARTICLE)', stArt);
        GPIEDBASE.CellValues [ColArticle, ARow] := TobP.GetValue ('(ARTICLE)');
        end;

    if stChamp <> '' then
        begin
        ValideSaisie (stChamp, ACol, ARow);
        end;
    CalculTTCSaisie;
end;

procedure TOF_CONTROLEFACTURE.GPIEDBASE_OnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var TobP : TOB;
begin
    if Action <> taConsult then
    begin
        TobP := TOB(GPIEDBASE.Objects [0, Ou]);
        if TobP <> nil then
        begin
            if not TobP.FieldExists ('ENPLUS') then
            begin
                GPIEDBASE.ColEditables [ColCategorie] := False;
                GPIEDBASE.ColEditables [ColFamille] := False;
            end;
        end;
    end;
end;

procedure TOF_CONTROLEFACTURE.GPIEDBASE_OnRowExit(Sender : TObject; Ou : Integer;
                                          var Cancel : Boolean; Chg : Boolean);
begin
if Action <> taConsult then
    begin
    ValideLigne(Ou);
    CalculTTCSaisie;
    end;
GPIEDBASE.ColEditables [ColCategorie] := True;
GPIEDBASE.ColEditables [ColFamille] := True;
end;

procedure TOF_CONTROLEFACTURE.GPIEDBASE_OnElipsisClick (Sender : TObject);
var TobP : Tob;
    stWhere : string;
begin
{if GPIEDBASE.Col = ColCategorie then
    LookUpList (GPIEDBASE, 'Catégorie de Taxe', 'CHOIXCOD', 'CC_CODE', 'CC_LIBELLE', SelectionCategorieTaxe, '',
                True, 1)
else }
if (GPIEDBASE.Col = ColFamille) and
   (GPIEDBASE.CellValues [ColCategorie, GPIEDBASE.Row] <> '') then
    LookUpList (GPIEDBASE, 'Régime de Taxe', 'CHOIXCOD', 'CC_CODE', 'CC_LIBELLE', SelectionFamilleTaxe, '',
                True, 0)
else if (GPIEDBASE.Col = ColArticle) then
    begin
    TobP := Tob(GPIEDBASE.Objects [0, GPIEDBASE.Row]);
    stWhere := ' AND GA_FAMILLETAXE' + Copy (TobP.GetValue ('GPB_CATEGORIETAXE'), 3, 1) +
                '="' + TobP.GetValue ('GPB_FAMILLETAXE') + '"';
    if GetArticleRecherche (GPIEDBASE, 'Recherche Article', '', '', '', stWhere) then
        begin
        TobP.PutValue ('GA_ARTICLE', GPIEDBASE.CellValues [ColArticle, GPIEDBASE.Row]);
        TobP.PutValue ('(ARTICLE)', GPIEDBASE.CellValues [ColArticle, GPIEDBASE.Row]);
        end;
    end;
bSaisie := True;
end;

Procedure TOF_CONTROLEFACTURE.GPIEDBASE_GetCellCanvas (Acol, ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState);
var TobP : TOB;
begin
TobP := TOB(GPIEDBASE.Objects [0, ARow]);
if (ACol = ColBaseSaisie) or (ACol = ColValeurSaisie) then
    Canvas.Font.Style := [fsBold]
else if TobP <> nil then
    begin
    if ((Acol = ColCategorie) or (ACol = ColFamille) or (ACol = ColArticle)) and
       (TobP.FieldExists ('ENPLUS')) then
        Canvas.Font.Style := [fsBold];
    end;
end;

procedure TOF_CONTROLEFACTURE.TTCSAISIE_OnChange (Sender : TObject);
begin
bSaisie := true;
end;

{==============================================================================================}
{==================================== Gestion des boutons =====================================}
{==============================================================================================}
procedure TOF_CONTROLEFACTURE.CopierClick (Sender: TObject);
var TobP : TOB;
begin
if Action <> taConsult then
    begin
    TobP := TOB(GPIEDBASE.Objects [0, GPIEDBASE.Row]);
    TobP.PutValue ('(BASESAISIE)', TobP.GetValue ('GPB_BASEDEV'));
    TobP.PutValue ('(VALEURSAISIE)', TobP.GetValue ('GPB_VALEURDEV'));
    GPIEDBASE.CellValues [ColBaseSaisie, GPIEDBASE.Row] := TobP.GetValue ('(BASESAISIE)');
    GPIEDBASE.CellValues [ColValeurSaisie, GPIEDBASE.Row] := TobP.GetValue ('(VALEURSAISIE)');
    CalculTTCSaisie;
    bSaisie := True;
    end;
end;

procedure TOF_CONTROLEFACTURE.CopierTTCClick (Sender: TObject);
begin
if Action <> taConsult then
    begin
    CalculTTCSaisie;
    SetControlText ('TTCSAISIE', GetControlText ('TTCSAISIECALC'));
    bSaisie := True;
    end;
end;

procedure TOF_CONTROLEFACTURE.ValiderClick (Sender: TObject);
begin
if Action <> taConsult then
    begin
    ValideLigne(GPIEDBASE.Row);
    dTotalTTCSaisie := Valeur (GetControlText ('TTCSAISIE'));
    stRefInterne := GetControlText ('REFINTERNE');
    bSaisie := False;
    bOk := true;
    end;
end;

procedure TOF_CONTROLEFACTURE.NouveauClick (Sender: TObject);
var TobP : TOB;
begin
GPIEDBASE.RowCount := GPIEDBASE.RowCount + 1;
TobP := Tob.Create ('PIEDBASE', TobPiedBases, -1);
TobP.AddChampSupValeur ('(BASESAISIE)', 0, false);
TobP.AddChampSupValeur ('(VALEURSAISIE)', 0, false);
TobP.AddChampSup ('ENPLUS', false);
GPIEDBASE.Objects [0, GPIEDBASE.RowCount - 1] := TobP;
bSaisie := true;
end;

procedure TOF_CONTROLEFACTURE.SupprimeClick (Sender: TObject);
var TobP : TOB;
begin
TobP := TOB(GPIEDBASE.Objects [0, GPIEDBASE.Row]);
if TobP.FieldExists ('ENPLUS') then
    begin
    TobP := TobPiedBases.FindFirst (['GPB_CATEGORIETAXE', 'GPB_FAMILLETAXE'],
                                    [GPIEDBASE.CellValues [ColCategorie, GPIEDBASE.Row],
                                     GPIEDBASE.CellValues [ColFamille, GPIEDBASE.Row]],
                                    false);
    if TobP <> nil then TobP.Free;
    GPIEDBASE.VidePile (False);
    AfficheGrille;
    end else
    begin
    TobP.PutValue ('(BASESAISIE)', 0);
    GPIEDBASE.CellValues [ColBaseSaisie, GPIEDBASE.Row] := '0';
    TobP.PutValue ('(VALEURSAISIE)', 0);
    GPIEDBASE.CellValues [ColValeurSaisie, GPIEDBASE.Row] := '0';
    end;
bSaisie := True;
end;

{==============================================================================================}
{==================================== Gestion =================================================}
{==============================================================================================}
procedure TOF_CONTROLEFACTURE.CalculTTCSaisie;
var dTTC : double;
    iRow : integer;
begin
dTTC := 0;
for iRow := 1 to GPIEDBASE.RowCount - 1 do
    begin
    dTTC := dTTC + Valeur (GPIEDBASE.CellValues [ColBaseSaisie, iRow]) +
                    Valeur (GPIEDBASE.CellValues [ColValeurSaisie, iRow]);
    end;
SetControlText ('TTCSAISIECALC', FloatToStr (dTTC));
end;

procedure TOF_CONTROLEFACTURE.AfficheGrille;
var TobP : TOB;
begin
if TobPiedBases.Detail.Count = 0 then
    begin
    GPIEDBASE.RowCount := 2;
    TobP := Tob.Create ('PIEDBASE', TobPiedBases, -1);
    TobP.AddChampSupValeur ('(BASESAISIE)', 0, false);
    TobP.AddChampSupValeur ('(VALEURSAISIE)', 0, false);
    TobP.AddChampSup ('ENPLUS', false);
    GPIEDBASE.objects [0, 1] := TobP;
    end else
    begin
    GPIEDBASE.RowCount := TobPiedBases.Detail.Count + 1;
    TobPiedBases.PutGridDetail (GPIEDBASE, false, false, GPIEDBASE.Titres [0]);
    end;
end;

{function TOF_CONTROLEFACTURE.SelectionCategorieTaxe : string;
begin
Result := 'CC_TYPE="GCX"';
end;}

function TOF_CONTROLEFACTURE.SelectionFamilleTaxe : string;
var iRow : integer;
    stPlus : string;
    TobP : TOB;
begin
Result := '';
TobP := TOB(GPIEDBASE.Objects [0, GPIEDBASE.Row]);
if TobP.FieldExists ('ENPLUS') then
    begin
    for iRow := 1 to GPIEDBASE.RowCount - 1 do
        begin
        if (iRow <> GPIEDBASE.Row) and
           (GPIEDBASE.CellValues [ColCategorie, iRow] = GPIEDBASE.CellValues [ColCategorie, GPIEDBASE.Row]) then
            begin
            TobP := TOB(GPIEDBASE.Objects [0, iRow]);
            if iRow > 1 then stPlus := stPlus + '","';
            stPlus := stPlus + TobP.GetValue ('GPB_FAMILLETAXE');
            end;
        end;
    Result := 'CC_TYPE="' + GPIEDBASE.CellValues [ColCategorie, GPIEDBASE.Row] + '"';
    if stPlus <> '' then Result := Result + ' AND CC_CODE NOT IN ("' + stPlus + '")';
    end;
end;

procedure TOF_CONTROLEFACTURE.TestCellule;
var TobP : TOB;
begin
TobP := TOB(GPIEDBASE.Objects [0, GPIEDBASE.Row]);
if (not TobP.FieldExists ('ENPLUS')) and
   (GPIEDBASE.Col <> ColBaseSaisie) and
   (GPIEDBASE.Col <> ColValeurSaisie) then
    begin
    GPIEDBASE.Col := ColBaseSaisie;
    end else if (TobP.FieldExists ('ENPLUS')) and
        (GPIEDBASE.Col <> ColBaseSaisie) and
        (GPIEDBASE.Col <> ColValeurSaisie) and
        (GPIEDBASE.Col <> ColCategorie) and
        (GPIEDBASE.Col <> ColFamille) and
        ((GPIEDBASE.Col <> ColArticle) or (not bArticle)) then
    begin
    GPIEDBASE.Col := ColCategorie;
    end;
if //(GPIEDBASE.Col = ColCategorie) or
   (GPIEDBASE.Col = ColFamille) or
   (GPIEDBASE.Col = ColArticle) then
    begin
    GPIEDBASE.ElipsisButton := True;
    end else
    begin
    GPIEDBASE.ElipsisButton := False;
    end;
end;

procedure TOF_CONTROLEFACTURE.ValideLigne(Ou : integer);
begin
ValideSaisie ('(BASESAISIE)', ColBaseSaisie, Ou);
ValideSaisie ('(VALEURSAISIE)', ColValeurSaisie, Ou);
end;

procedure TOF_CONTROLEFACTURE.ValideSaisie (stChamp : string; iCol, iRow : integer);
var dSaisie : double;
    TobPied : TOB;
begin
TobPied := TOB(GPIEDBASE.Objects [0, iRow]);
dSaisie := Valeur(GPIEDBASE.Cellvalues[iCol, iRow]);
if Valeur (TobPied.GetValue (stChamp)) <> dSaisie then bSaisie := True;
TobPied.PutValue (stChamp, dSaisie);
end;

Initialization
  registerclasses ( [ TOF_CONTROLEFACTURE ] ) ;

end.

