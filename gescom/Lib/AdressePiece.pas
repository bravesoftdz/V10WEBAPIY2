unit AdressePiece;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Hctrls, UTOB, HPanel, HEnt1, UIUtil, HSysMenu, SaisUtil,
  Buttons, AglInit,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
  Math, FactAdresse,EntGC, TntButtons, TntGrids ;

type
  TFAdrPiece = class(TForm)
    G_ADR: THGrid;
    Label1: TLabel;
    Label2: TLabel;
    bNouveau: THBitBtn;
    G_LIG: THGrid;
    bValider: THBitBtn;
    bAnnuler: THBitBtn;
    procedure FormShow(Sender: TObject);
    procedure bNouveauClick(Sender: TObject);
    procedure G_LIGDblClick(Sender: TObject);
    procedure G_ADRClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure bValiderClick(Sender: TObject);
  private
    { Déclarations privées }
    LesColsLIG, LesColsADR : string;
    procedure TailleColonnes(Grid : THGrid);
    procedure DessineCell ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
  public
    { Déclarations publiques }
  end;

var
  FAdrPiece: TFAdrPiece;
  lValide : boolean;
  TobAdr, TobLigGen, TobLig, TobTemp, TobTemp2, ToBAdrFact : TOB;
  LivFac, i_NumPiece : integer;
  CodeTiers, CodeTiersLivre, RefPiece : string;

Const St_Sel = 'þ';

Procedure EntreeAdressePiece (TobPiece, TobAdresses : TOB; Option : Integer);
Procedure ValideAdressePiece (TobPiece, TobAdresses : TOB);

implementation

{$R *.DFM}

procedure EntreeAdressePiece (TobPiece, TobAdresses : TOB; Option : Integer);
var FF : TFAdrPiece ;
    PPANEL  : THPanel ;
    i_ind1, i_ind2 : Integer;
    TobTemp : TOB;
begin
FF := TFAdrPiece.Create(Application) ;
//  Préparation des données de la Forme
lValide := False;
FF.bValider.Enabled := False;
//FF.bAnnuler.Enabled := False;
//  Duplication des tob piece et adresses dans des tobs locales
TobLig := TOB.Create('', nil, -1);
TobLigGen := TOB.Create('', nil, -1);
LivFac := Option;
i_NumPiece := TobPiece.GetValue('GP_NUMERO');
CodeTiers := TobPiece.GetValue('GP_TIERS');
CodeTiersLivre := TobPiece.GetValue('GP_TIERSLIVRE');
RefPiece:=TOBPiece.GetValue('GP_NATUREPIECEG')+';'+TOBPiece.GetValue('GP_SOUCHE')+';'+IntToStr(TOBPiece.GetValue('GP_NUMERO'))+';' ;

for i_ind1 := 0 to TobPiece.Detail.Count - 1 do
    begin
    TobTemp := TOB.Create('LIGNE', TobLigGen, -1);
    TobTemp.Dupliquer(TobPiece.Detail[i_ind1] , True, True);
    if TobPiece.Detail[i_ind1].GetValue('GL_ARTICLE')='' then  continue;
    TobTemp2 := TOB.Create('LIGNE', TobLig, -1);
    TobTemp2.Dupliquer(TobPiece.Detail[i_ind1] , True, True);
    end;

if TobLig.Detail.Count<=0 then
   begin
   TobLig.Free;
   TobLigGen.Free;
   TobLig := Nil;
   TobLigGen := Nil;
   exit;
   end;

TobAdrFact := TOB.Create('ADRESSES', nil, -1);
TobAdr := TOB.Create('ADRESSES', nil, -1);

for i_ind1 := 0 to TobAdresses.Detail.Count - 1 do
    begin
//  On ecarte l'adresse de facturation de la Grid
    if i_ind1 = 1 then
        begin
        TobAdrFact.Dupliquer(TobAdresses.Detail[i_ind1] , True, True);
        Continue;
        end;
    TobTemp := TOB.Create('ADRESSES', TobAdr, -1);
    TobTemp.Dupliquer(TobAdresses.Detail[i_ind1] , True, True);
    end;
if Option = 0 then
    FF.Caption := 'Adresses de livraison de la pièce '
else
    FF.Caption := 'Adresses de facturation de la pièce ';
//
SourisSablier;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside (FF, PPANEL) ;
   FF.Show ;
   END ;
if lValide then
    begin
    for i_ind1 := 0 to TobLig.Detail.Count - 1 do
        begin
        TobTemp := TobLigGen.FindFirst(['GL_NUMLIGNE'],[TobLig.Detail[i_ind1].GetValue('GL_NUMLIGNE')], False);
        if TobTemp <> nil then TobTemp.PutValue('GL_NUMADRESSELIVR',TobLig.Detail[i_ind1].GetValue('GL_NUMADRESSELIVR'));
        end;
    for i_ind1 := 0 to TobLigGen.Detail.Count - 1 do
        TobPiece.Detail[i_ind1].PutValue('GL_NUMADRESSELIVR',TobLigGen.Detail[i_ind1].GetValue('GL_NUMADRESSELIVR'));
    for i_ind1 := 0 to TobAdr.Detail.Count - 1 do
        begin
//  On ajoute l'adresse de facturation precedement ecarte de la Grid
        if i_ind1 = 0 then
            begin
            TobAdresses.Detail[1].Dupliquer(TobAdrFact , True, True);
            TobAdresses.Detail[1].SetAllModifie(False);
            end;
        i_ind2 := i_ind1;
        if i_ind1 > 0 then inc(i_ind2);
        if TobAdr.Detail.Count + 1 > TobAdresses.Detail.Count then
          TOB.Create('ADRESSES', TobAdresses, -1);
        TobAdresses.Detail[i_ind2].Dupliquer(TobAdr.Detail[i_ind1] , True, True);
        end;
    end;
TobLig.Free; TobLig := Nil;
TobLigGen.Free; TobLigGen := Nil;
TobAdr.Free; TobAdr := Nil;
END ;

//  Validation des adresses pieces eventuellement crees, hors adresses de livraison
//  et de facturation de l'entete.
procedure ValideAdressePiece (TobPiece, TobAdresses : TOB);
var
    i_ind1, i_ind2, i_adrnew, i_adrold : Integer;
    lExiste : boolean;
    TobTemp : TOB;
    Q : TQuery;
    Select : string;
begin
  if not VH_GC.GCIfDefCEGID then
  begin
    for i_ind1 := 0 to TobAdresses.Detail.Count - 1 do
        begin
            TobTemp := TobAdresses.Detail[i_ind1];
            i_adrold := TobTemp.GetValue('ADR_NUMEROADRESSE');
    //  on elimine les lignes à adresses > -2 (deja créees ou affectées à l'entête
            if i_adrold >= -2 then Continue;
            i_adrnew := CreerAdresse(TobTemp);
            for i_ind2 := 0 to TobPiece.Detail.Count - 1 do
                if TobPiece.Detail[i_ind2].GetValue('GL_NUMADRESSELIVR') = i_adrold then
                    TobPiece.Detail[i_ind2].PutValue('GL_NUMADRESSELIVR', i_adrnew);
        end;
    //  on recherche toutes les adresses utilisées par les lignes du document
    TobTemp := TOB.Create('',nil,-1);
    Select := 'SELECT ADR_NUMEROADRESSE FROM ADRESSES WHERE ADR_NUMEROADRESSE IN ' +
              '(SELECT DISTINCT GL_NUMADRESSELIVR FROM LIGNE WHERE ' +
              'GL_NATUREPIECEG = "' + string(TobPiece.GetValue('GP_NATUREPIECEG')) + '" AND ' +
              'GL_DATEPIECE = "' + USDATETIME(TobPiece.GetValue('GP_DATEPIECE')) + '" AND ' +
              'GL_SOUCHE = "' + string(TobPiece.GetValue('GP_SOUCHE')) + '" AND ' +
              'GL_NUMERO = ' + IntToStr(TobPiece.GetValue('GP_NUMERO')) + ' AND ' +
              'GL_INDICEG = "' + string(TobPiece.GetValue('GP_INDICEG')) + '")';
    Q:=OpenSQL(Select, True,-1,'',true);
    TobTemp.LoadDetailDB('ADRESSES', '', '', Q, True);
    Ferme(Q) ;
    //  on elimine toutes les adresses de la piece non utilisées dans les lignes
    for i_ind1 := 0 to TobAdresses.Detail.Count - 1 do
        begin
        if i_ind1 = 1 then Continue;  // on ne traite pas l'adresse de facturation...
        i_adrold := TobAdresses.Detail[i_ind1].GetValue('ADR_NUMEROADRESSE');
        lExiste := False;
        for i_ind2 := 0 to TobTemp.Detail.Count - 1 do
            if TobTemp.Detail[i_ind2].GetValue('ADR_NUMEROADRESSE') = i_adrold then
                begin
                lExiste := True;
                Break;
                end;
        if (not lExiste) and (TobTemp.Detail.Count <> 0) then
            TobAdresses.Detail[i_ind1].DeleteDB;
        end;
  end;
end;
//
//***************************************************************************
//
procedure TFAdrPiece.FormShow(Sender: TObject);
var
    st_temp : string;
    i_adrsel, i_ind1 : integer;

begin
//  Chargement des HGrid avec les données retournées par les listes
G_ADR.ListeParam := 'GCADRPIECE';
G_LIG.ListeParam := 'GCLIGPIECE';
AffecteGrid(G_ADR,taConsult);
AffecteGrid(G_LIG,taConsult);
G_LIG.GetCellCanvas := DessineCell;
st_temp := Copy(G_ADR.Titres.Strings[0], 0, Length(G_ADR.Titres.Strings[0]) - 1);
TobAdr.PutGridDetail(G_ADR, False, True, st_temp);
st_temp := Copy(G_LIG.Titres.Strings[0], 0, Length(G_LIG.Titres.Strings[0]) - 1);
//TobLig.PutGridDetail(G_LIG, True, True, st_temp);
TobLig.PutGridDetail(G_LIG, False, True, st_temp);
//  Positionnement des lignes correspondant à l'adresse 1
i_adrsel := StrToInt(G_ADR.Cells[0, 1]);
for i_ind1 := 0 to G_LIG.RowCount - 2 do
    if TobLig.Detail[i_ind1].GetValue('GL_NUMADRESSELIVR') = i_adrsel then
        BEGIN
        G_LIG.Cells[G_LIG.ColCount - 1, i_ind1 + 1] := '*';
        G_LIG.Cells[1, i_ind1 + 1] := St_Sel;
        END else
        BEGIN
        G_LIG.Cells[G_LIG.ColCount - 1, i_ind1 + 1] := ' ';
        G_LIG.Cells[1, i_ind1 + 1] := ' ';
        END;
//  Memorisation des titres de colonnes (noms des champs)
LesColsLIG := G_LIG.Titres[0];
LesColsADR := G_ADR.Titres[0];
//  Mise à 0 de la largeur des colonnes numeros et de colonne
G_ADR.ColWidths[0] := 0;
G_LIG.ColWidths[0] := 0;
G_LIG.ColWidths[G_LIG.ColCount - 1] := 0;
G_LIG.ColWidths[G_LIG.ColCount - 2] := 0;
//  Gestion des largeurs de colonnes
TailleColonnes(G_ADR);
TailleColonnes(G_LIG);
//
end;

procedure TFAdrPiece.G_ADRClick(Sender: TObject);
var
    i_ind1, i_adrsel : integer;
begin
//  Rafraichissement de la HGrid lignes
TobLig.PutGridDetail(G_LIG, False, True, LesColsLIG);
//  Recherche des lignes piece avec adresse de livraison egale à l'adresse selectionnée
i_adrsel := StrToInt(G_ADR.Cells[0, G_ADR.Row]);
for i_ind1 := 0 to G_LIG.RowCount - 2 do
    if TobLig.Detail[i_ind1].GetValue('GL_NUMADRESSELIVR') = i_adrsel then
        BEGIN
        G_LIG.Cells[G_LIG.ColCount - 1, i_ind1 + 1] := '*';
        G_LIG.Cells[1, i_ind1 + 1] := St_Sel;
        END else
        BEGIN
        G_LIG.Cells[G_LIG.ColCount - 1, i_ind1 + 1] := ' ';
        G_LIG.Cells[1, i_ind1 + 1] := ' ';
        END;
end;

procedure TFAdrPiece.G_LIGDblClick(Sender: TObject);
var i_adrnum, i_adrlig : Integer;
    Sel : boolean;
begin
//  Validation des boutons Valider et Annuler
bValider.Enabled := True;
//bAnnuler.Enabled := True;
Sel := False;
//  sur double clic, affectation ou desaffectation de l'adresse à la ligne
i_adrnum := TobAdr.Detail[G_ADR.Row - 1].GetValue('ADR_NUMEROADRESSE');
i_adrlig := TobLig.Detail[G_LIG.Row - 1].GetValue('GL_NUMADRESSELIVR');
if TobLig.Detail[G_LIG.Row - 1].GetValue('GL_NATUREPIECEG') <> '*' then
    begin
    TobLig.Detail[G_LIG.Row - 1].PutValue('GL_NATUREPIECEG', '*');
    Sel := True;
    end
    else
    begin
    if i_adrnum = i_adrlig then
        begin
        TobLig.Detail[G_LIG.Row - 1].PutValue('GL_NATUREPIECEG', ' ');
        Sel := False;
        i_adrnum := 0;
        end;
    end;
//  rafraichissement de la HGrid Lignes
TobLig.Detail[G_LIG.Row - 1].PutValue('GL_NUMADRESSELIVR', i_adrnum);
TobLig.Detail[G_LIG.Row - 1].PutLigneGrid(G_LIG, G_LIG.Row, False, False, LesColsLIG);
if Sel then G_LIG.Cells[1, G_LIG.Row] := St_Sel else G_LIG.Cells[1, G_LIG.Row] := ' ';
end;

procedure TFAdrPiece.bNouveauClick(Sender: TObject);
var
    Valeur, Select : string;
    TobTemp : TOB;
begin
Select := '((ADR_REFCODE = "' + CodeTiers +
          '" OR ADR_REFCODE = "' + CodeTiersLivre +
          '") AND ADR_TYPEADRESSE = "TIE") OR (ADR_REFCODE = "' + RefPiece +
          '" AND ADR_TYPEADRESSE = "PIE")' ;
Valeur := AGLLanceFiche('GC','GCADRESSES_MUL', '', '', Select);
if Valeur <> '' then
    begin
    TobTemp := TOB.Create('ADRESSES', TobAdr, -1);
    TobTemp.PutValue('ADR_NUMEROADRESSE', Valeur);
    TobTemp.PutValue('ADR_TYPEADRESSE', 'TIE');
    TobTemp.ChargeCle1;
    TobTemp.LoadDB(False);
//  Mise du numero d'adresse en negatif pour gerer la nouvelle adresse
    TobTemp.PutValue('ADR_NUMEROADRESSE', -(G_ADR.RowCount + 1));
    //TobTemp.PutValue('ADR_REFCODE', ReadTokenSt(Valeur) + ';' + ReadTokenSt(Valeur) + ';' +
     //                               IntToStr(i_NumPiece) + ';');
    G_ADR.InsertRow(G_ADR.RowCount);
    TobTemp.PutLigneGrid(G_ADR, G_ADR.RowCount - 1, False, False, LesColsADR);
    TailleColonnes(G_ADR);
    end;
end;

procedure TFAdrPiece.bValiderClick(Sender: TObject);
begin
lValide := True;
Close;
end;

procedure TFAdrPiece.bAnnulerClick(Sender: TObject);
begin
Close;
end;

procedure TFAdrPiece.TailleColonnes(Grid : THGrid);
var
    LabelTemp1 : TLabel;
    i_ind1, i_ind2, LMax : integer;
begin
//  Gestion des largeurs de colonnes sauf celles à largeur initialisées à 0
LabelTemp1 := TLabel.Create(Grid);
LabelTemp1.AutoSize := True;
LabelTemp1.Font := Grid.Font;
for i_ind1 := 0 to Grid.ColCount - 1 do
begin
    if Grid.ColWidths[i_ind1] = 0 then Continue;
    LabelTemp1.Caption := Grid.Cells[i_ind1, 0];
    LMax := LabelTemp1.Width;
    for i_ind2 := 1 to Grid.RowCount - 1 do
    begin
        LabelTemp1.Caption := Grid.Cells[i_ind1, i_ind2];
        LMax := Max(LMax, LabelTemp1.Width);
    end;
    Grid.ColWidths[i_ind1] := LMax + (Length(Grid.Cells[i_ind1, 0]) * 2);
end;
LabelTemp1.Free;
end;

procedure TFAdrPiece.DessineCell ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
Var Coord : TRect;
    ST : string;
BEGIN
if (ACol = 1) AND (ARow > 0) then
    BEGIN
    Coord := G_LIG.CellRect (ACol, ARow);
    Canvas.Font.Name:='Wingdings' ;
    Canvas.Font.Size:=10 ;
    Canvas.Font.Color:=TColor ($0000FF00) ; // Vert pleine intensité
    Canvas.Font.Style := [fsBold];
    st := G_LIG.Cells [ACol, Arow];
    Canvas.TextOut ( (Coord.Left+Coord.Right) div 2 - Canvas.TextWidth(st) div 2,
                     (Coord.Top+ Coord.Bottom) div 2 - Canvas.TextHeight(st) div 2, st);
    END;
END;


end.
