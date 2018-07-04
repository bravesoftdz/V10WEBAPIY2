{ Unité : Source TOF de la fiche CPBONSAPAYER
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001   14/03/06    JP     Création de l'unité
--------------------------------------------------------------------------------------}
unit CPBONSAPAYER_TOF;

interface

uses
  Controls, Classes, Vierge,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  Grids, SysUtils, HCtrls, UTob, HEnt1, UTOF;

type
  TOF_CPBONSAPAYER = class(TOF)
    FListe   : THGrid;

    procedure OnArgument(S : string); override;
  private
    TobBap : TOB;

    procedure AfficheBap(Clef : string);
    procedure DessineGrille;
    procedure ListeDbClick(Sender : TObject);
  end;

procedure CpLanceFiche_FicheBap(Clef : string);

implementation

uses
  uLibBonAPayer, ULibPieceCompta, ULibEcriture, CPSAISIEPIECE_TOF;

const
  COL_JOURNAL       = 0;
  COL_DATECOMPTABLE = 1;
  COL_NUMEROPIECE   = 2;
  COL_CIRCUITBAP    = 3;
  COL_VISEUR        = 4;
  COL_VISEUR1       = 5;
  COL_VISEUR2       = 6;
  COL_RELANCE1      = 7;
  COL_ECHEANCEBAP   = 8;
  COL_STATUTBAP     = 9;
  COL_EXERCICE      = 10;
  COL_CODEJOURNAL   = 11;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_FicheBap(Clef : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPBONSAPAYER', '', '', Clef + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBONSAPAYER.AfficheBap(Clef : string);
{---------------------------------------------------------------------------------------}
var
  RecC : TClefPiece;
begin
  RecC := DecodeClefTP(Clef, ',');
  RecC.E_QUALIFPIECE := 'N';
  TobBap := TOB.Create('µµµ', nil, -1);
  try
    TobBap.LoadDetailFromSQL('SELECT * FROM CPBONSAPAYER WHERE BAP_JOURNAL = "' + RecC.E_JOURNAL + '" AND ' +
                             'BAP_EXERCICE = "' + RecC.E_EXERCICE + '" AND ' +
                             'BAP_DATECOMPTABLE = "' + USDateTime(RecC.E_DATECOMPTABLE) + '" AND ' +
                             'BAP_NUMEROPIECE = ' + IntToStr(RecC.E_NUMEROPIECE));
    {2 fois le journal afin de mémoriser le code journal, après le RechDom sur la colonne COL_JOURNAL
     afin de pouvoir lancer la saisie paramétrable}
    TobBap.PutGridDetail(FListe, False, False, 'BAP_JOURNAL;BAP_DATECOMPTABLE;BAP_NUMEROPIECE;BAP_CIRCUITBAP;BAP_VISEUR;BAP_VISEUR1;BAP_VISEUR2;BAP_RELANCE1;BAP_ECHEANCEBAP;BAP_STATUTBAP;BAP_EXERCICE;BAP_JOURNAL;');
    DessineGrille;
  finally
    if Assigned(TobBap) then FreeAndNil(TobBap);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBONSAPAYER.DessineGrille;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  FListe.ColAligns[COL_RELANCE1     ] := taCenter;
  FListe.ColAligns[COL_DATECOMPTABLE] := taCenter;
  FListe.ColAligns[COL_NUMEROPIECE  ] := taRightJustify;
  FListe.ColWidths[COL_EXERCICE     ] := -1;
  {Pour mémoriser le code journal, après le RechDom sur la colonne COL_JOURNAL afin
   de pouvoir lancer la saisie paramétrable}
  FListe.ColWidths[COL_CODEJOURNAL  ] := -1;

  for n := 1 to FListe.RowCount - 1 do begin
    FListe.Cells[COL_JOURNAL   , n] := RechDom('CPJOURNAL'    , FListe.Cells[COL_JOURNAL   , n], False);
    FListe.Cells[COL_CIRCUITBAP, n] := RechDom('CPCIRCUITBAP' , FListe.Cells[COL_CIRCUITBAP, n], False);
    FListe.Cells[COL_VISEUR    , n] := RechDom('TTUTILISATEUR', FListe.Cells[COL_VISEUR    , n], False);
    FListe.Cells[COL_VISEUR1   , n] := RechDom('TTUTILISATEUR', FListe.Cells[COL_VISEUR1   , n], False);
    FListe.Cells[COL_VISEUR2   , n] := RechDom('TTUTILISATEUR', FListe.Cells[COL_VISEUR2   , n], False);
    FListe.Cells[COL_STATUTBAP , n] := RechDom('CPSTATUTBAP'  , FListe.Cells[COL_STATUTBAP , n], False);
    if FListe.Cells[COL_RELANCE1, n] = 'X' then FListe.Cells[COL_RELANCE1, n] := 'OUI'
                                           else FListe.Cells[COL_RELANCE1, n] := 'NON';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBONSAPAYER.ListeDbClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  lInfoEcr  : TInfoEcriture;
  lPieceCpt : TPieceCompta;
begin
  // Chargement pièce
  lInfoEcr  := TInfoEcriture.Create(V_PGI.SchemaName);
  lPieceCpt := TPieceCompta.CreerPiece(lInfoEcr);
  try
    // -- Champs obligatoires...
    lPieceCpt.PutEntete('E_JOURNAL',       FListe.Cells[COL_CODEJOURNAL, FListe.Row]);
    lPieceCpt.PutEntete('E_EXERCICE',      FListe.Cells[COL_EXERCICE   , FListe.Row]);
    lPieceCpt.PutEntete('E_DATECOMPTABLE', StrToDate(FListe.Cells[COL_DATECOMPTABLE, FListe.Row]));
    lPieceCpt.PutEntete('E_NUMEROPIECE',   ValeurI(FListe.Cells[COL_NUMEROPIECE, FListe.Row]));
    lPieceCpt.PutEntete('E_QUALIFPIECE',   'N');

    // -- Chargement
    lPieceCpt.LoadFromSQL ;

    // Ouverture Saisie
    ModifiePieceCompta(lPieceCpt, taConsult);
  finally
    FreeAndNil(lInfoEcr);
    FreeAndNil(lPieceCpt);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPBONSAPAYER.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  inherited;
  FListe := THGrid(GetControl('FLISTE'));
  {Récupération de la Clef}
  ch := ReadTokenSt(S);
  AfficheBap(ch);
  FListe.OnDblClick := ListeDbClick;
  Ecran.HelpContext := 7509076;
end;

initialization
  RegisterClasses([TOF_CPBONSAPAYER]);

end.
 