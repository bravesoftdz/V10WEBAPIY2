{ Unité : Source TOF de la fiche TRLISTEECRCOMPTA
  Objet : Affichage des écritures à intégrer dans la compta, en présentation comptables
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            24/07/03    JP     Création de l'unité
 0.91            16/09/03    JP     Gestion de la TVA
 6.0x.xxx.xxx    09/08/04    JP     Suppression de la notion de devise qui ne servait à rien
 6.30.001.004    21/03/05    JP     Ajout d'un test sur l'équilibrage des écritures
 7.09.001.010    30/03/07    JP     FQ 10427 : Problème d'arrondi
--------------------------------------------------------------------------------------}

unit TRLISTEECRCOMPTA_TOF ;

interface

uses
  Windows, Controls, Classes, Graphics,
  {$IFDEF EAGLCLIENT}
    MaineAGL,
  {$ELSE}
    FE_Main, PrintDBG, HTB97,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, UTOF, ExtCtrls, UTOB, Grids;

type
  TOF_TRLISTEECRCOMPTA = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnCancel                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnClose                  ; override ;
  private
    ListeResultat : Boolean;
    Devise        : string;
    FListe        : THGrid;
    ModePiece     : Boolean ;

    procedure FListeDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
    {$IFNDEF EAGLCLIENT}
    procedure Imprimer      (Sender : TObject);
    {$ENDIF}
    procedure AffDetail     (Sender : TObject);
    procedure ChargerGrille;
    procedure ChargerResult;
    procedure DessinerGrille;

    // Ajout du mode piece
    procedure ChargerGrillePiece ;
    procedure ChargerResultPiece ;

  end ;

// Version avce TPieceCompta
function  TRAfficheListePieces( Liste : TOB ) : Boolean;
procedure TRAffichePiecesInteg( Liste : TOB ) ;

implementation

uses
  Commun, TRLIGNEECRCOMPTA_TOF, Constantes, UProcEcriture;

var
  TobGenerale : TOB;
  Resultat    : Boolean;

{---------------------------------------------------------------------------------------}
function  TRAfficheListePieces( Liste : TOB ) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  TobGenerale := Liste;
  AGLLanceFiche('TR', 'TRLISTEECRCOMPTA', '', '', 'NON;PIECE;');
  Result := Resultat;
end;

{---------------------------------------------------------------------------------------}
procedure TRAffichePiecesInteg( Liste : TOB ) ;
{---------------------------------------------------------------------------------------}
begin
  TobGenerale := Liste;
  AGLLanceFiche('TR', 'TRLISTEECRCOMPTA', '', '', 'OUI;PIECE;');
end;


{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;

  Resultat := False;

  ListeResultat := ReadTokenSt(S) = 'OUI';
  ModePiece     := ReadTokenSt(S) = 'PIECE';

  {$IFDEF EAGLCLIENT}
  SetControlVisible('BIMPRIMER', False);
  {$ELSE}
  SetControlVisible('BIMPRIMER', True);
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := Imprimer;
  {$ENDIF}

  AssignDrapeau(TImage(GetControl('IDEV')), Devise);

  {Dans la liste des résultat, pas de devise}
  if ListeResultat then
    SetControlCaption('LBDEVISE', '');

  FListe := THGrid(GetControl('FLISTE'));
  if ListeResultat then begin
    FListe.OnDrawCell := FListeDrawCell;
    Ecran.Caption := 'Résultat de l''injection en comptabilité';
    UpdateCaption(Ecran);
  end
  else
    FListe.OnDblClick := AffDetail; {Pour l'affichage du détail d'une ligne}

  if ModePiece then begin
    if ListeResultat then ChargerResultPiece {Chargement du résultat de l'intégration dans la grille}
                     else ChargerGrillePiece;{Chargement des valeurs à injecter dans la comptabilité}
  end else begin
    if ListeResultat then ChargerResult {Chargement du résultat de l'intégration dans la grille}
                     else ChargerGrille;{Chargement des valeurs à injecter dans la comptabilité}
  end ;

end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.OnCancel;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  {True pour interrompre l'intégration des écritures }
  Resultat := False;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  C : Double;
  D : Double;
  M : string;
begin
  inherited;
  C := 0;
  D := 0;
  {21/03/05 : Calcul du total du Crédit et du Débit}
  for n := 0 to TobGenerale.Detail.Count - 1 do begin
    {On teste le général, car c'est le test qui est appliqué partout pour savoir si la Tva est gérée}
    if TobGenerale.Detail[n].GetString('E_GENERAL') <> '' then begin
      D := Arrondi(TobGenerale.Detail[n].GetDouble('E_DEBIT'), V_PGI.OkDecV) + D;
      C := Arrondi(TobGenerale.Detail[n].GetDouble('E_CREDIT'), V_PGI.OkDecV) + C;
    end;
  end;

  {21/03/05 : On teste si l'ensemble des pièces est équilibré : c'est tout ou rien !
   30/03/07 : FQ 10427}
  if Arrondi(D - C, V_PGI.OkDecV) <> 0 then begin
    M := 'Certaines pièces sont déséquilibrées. Le traitement est interrompu.';
    if NatErr_Tva in ErreurCategorie.Ecriture then
      M := M + #13'Vérifiez notamment le paramétrage de la TVA en comptabilité.';

    HShowMessage('0;' + Ecran.Caption + ';' + M + ';W;O;O;O;', '', '');
    Resultat := False;
    {Réinitialisation des deux ensemble de gestion des erreurs puisque l'on abandonne le traitement}
    InitGestionErreur;
  end
  else
    {True pour lancer l'intégration des écritures proprement dite}
    Resultat := True;
end ;

{$IFDEF EAGLCLIENT}
// Créer l'état pour un LanceEtatTob
{$ELSE}
{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.Imprimer(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  PrintDBGrid(FListe, nil, 'Liste des écritures à intégrer', '');
end;
{$ENDIF}

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.ChargerGrille;
{---------------------------------------------------------------------------------------}
var
  k, n, p : Integer;
begin
  k := TobGenerale.Detail.Count;
  p := 0;
  for n := 0 to k - 1 do begin
    {JP 16/09/03 : Pour chaque écriture de tréso, on génère 3 écritures comptables.
                   Cependant, si la le flux n'est pas assujetti à TVA, la ligne TVA
                   est vide et ne sera pas injectée en compta => on ne l'affiche pas}
    if TobGenerale.Detail[n].GetValue('E_GENERAL') <> '' then begin
      FListe.Cells[1, p + 1] := TobGenerale.Detail[n].GetValue('E_JOURNAL');
      FListe.Cells[2, p + 1] := TobGenerale.Detail[n].GetValue('E_DATECOMPTABLE');
      FListe.Cells[3, p + 1] := TobGenerale.Detail[n].GetValue('E_GENERAL');
      FListe.Cells[4, p + 1] := TobGenerale.Detail[n].GetValue('E_REFINTERNE');
      FListe.Cells[5, p + 1] := FloatToStr(Arrondi(TobGenerale.Detail[n].GetValue('E_DEBIT'), 2));
      FListe.Cells[6, p + 1] := FloatToStr(Arrondi(TobGenerale.Detail[n].GetValue('E_CREDIT'), 2));
      FListe.Objects[0, p+1] := TobGenerale.Detail[n];
      Inc(p);
    end
    else begin
      {21/03/05 : Si le général est vide et que le numéro de ligne = 3, on est sur une ligne de Tva.
                  Vraisemblablement, le paramétrage de la Tva est incomplet en compta}
      if TobGenerale.Detail[n].GetInteger('E_NUMLIGNE') = 3 then
        ErreurCategorie.Ecriture := [NatErr_Tva];
    end;
  end;
  FListe.RowCount := p + 1;

  {24/05/05 : Centrage des champs Debit et Crédit sur la droite}
  FListe.ColAligns[5] := taRightJustify;
  FListe.ColAligns[6] := taRightJustify;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.ChargerResult;
{---------------------------------------------------------------------------------------}
var
  k, n, p : Integer;
begin
  {On commence par redessiner la grille}
  FListe.DefaultDrawing := False;
  DessinerGrille;
  p := 0;

  {Remplissage de la grille}
  k := TobGenerale.Detail.Count;
  FListe.RowCount := k + 1;
  for n := 0 to k - 1 do begin
    {JP 16/09/03 : Pour chaque écriture de tréso, on génère 3 écritures comptables.
                   Cependant, si la le flux n'est pas assujetti à TVA, la ligne TVA
                   est vide et ne sera pas injectée en compta => on ne l'affiche pas}
    if TobGenerale.Detail[n].GetValue('E_GENERAL') <> '' then begin
      FListe.Cells[1, p + 1] := TobGenerale.Detail[n].GetValue('E_JOURNAL');
      FListe.Cells[2, p + 1] := TobGenerale.Detail[n].GetValue('E_GENERAL');
      FListe.Cells[3, p + 1] := TobGenerale.Detail[n].GetValue('E_REFINTERNE');
      FListe.Cells[4, p + 1] := TobGenerale.Detail[n].GetValue('RESULTAT');
      FListe.Cells[5, p + 1] := TobGenerale.Detail[n].GetValue('COMMENTAIRE');
      Inc(p);
    end;
  end;
  FListe.RowCount := p + 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.DessinerGrille;
{---------------------------------------------------------------------------------------}
begin

  if not (ModePiece or ListeResultat) then Exit ;

  if ModePiece then begin

    if ListeResultat then begin
      FListe.ColCount := 6;
      FListe.Cells[0, 0] := TraduireMemoire('Dossier');
      FListe.Cells[1, 0] := TraduireMemoire('Journal');
      FListe.Cells[2, 0] := TraduireMemoire('Compte');
      FListe.Cells[3, 0] := TraduireMemoire('Libelle');
      FListe.Cells[4, 0] := TraduireMemoire('Résultat');
      FListe.Cells[5, 0] := TraduireMemoire('Commentaires');
                            (*
      FListe.Titres.Clear;
      FListe.Titres.Add(TraduireMemoire('Dossier'));
      FListe.Titres.Add(TraduireMemoire('Journal'));
      FListe.Titres.Add(TraduireMemoire('Compte'));
      FListe.Titres.Add(TraduireMemoire('Libelle'));
      FListe.Titres.Add(TraduireMemoire('Résultat'));
      FListe.Titres.Add(TraduireMemoire('Commentaires'));
      *)
      if IsTresoMultiSoc
        then FListe.ColWidths[0] := 70
        else FListe.ColWidths[0] := -1 ;
      FListe.ColWidths[1] := 50;
      FListe.ColWidths[2] := 100;
      FListe.ColWidths[3] := 135;
      FListe.ColWidths[4] := 50;
      FListe.ColWidths[5] := 305;
    end else begin
      FListe.ColCount := 7;
      FListe.Cells[0, 0] := TraduireMemoire('Dossier');
      FListe.Cells[1, 0] := TraduireMemoire('Journal');
      FListe.Cells[2, 0] := TraduireMemoire('Date Comptable');
      FListe.Cells[3, 0] := TraduireMemoire('Compte');
      FListe.Cells[4, 0] := TraduireMemoire('Référence');
      FListe.Cells[5, 0] := TraduireMemoire('Débit');
      FListe.Cells[6, 0] := TraduireMemoire('Crédit');
      (*
      FListe.Titres.Clear;
      FListe.Titres.Add(TraduireMemoire('Dossier'));
      FListe.Titres.Add(TraduireMemoire('Journal'));
      FListe.Titres.Add(TraduireMemoire('Date Comptable'));
      FListe.Titres.Add(TraduireMemoire('Compte'));
      FListe.Titres.Add(TraduireMemoire('Référence'));
      FListe.Titres.Add(TraduireMemoire('Débit'));
      FListe.Titres.Add(TraduireMemoire('Crédit'));
      *)
      if IsTresoMultiSoc
        then FListe.ColWidths[0] := 70
        else FListe.ColWidths[0] := -1 ;
      FListe.ColWidths[1] := 50;
      FListe.ColWidths[2] := 90;
      FListe.ColWidths[3] := 100;
      FListe.ColWidths[4] := 135;
      FListe.ColWidths[5] := 90;
      FListe.ColWidths[6] := 90;
    end ;
  end else begin
    FListe.ColCount := 5;
    FListe.ColWidths[1] := 50;
    FListe.ColWidths[2] := 100;
    FListe.ColWidths[3] := 125;
    FListe.ColWidths[4] := 50;
    FListe.ColWidths[5] := 305;
    FListe.Titres.Clear;
    FListe.Titres.Add(TraduireMemoire('Journal'));
    FListe.Titres.Add(TraduireMemoire('Compte'));
    FListe.Titres.Add(TraduireMemoire('Libelle'));
    FListe.Titres.Add(TraduireMemoire('Résultat'));
    FListe.Titres.Add(TraduireMemoire('Commentaires'));
  end ;

 // FListe.UpdateTitres;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.FListeDrawCell(Sender: TObject; Col, Row: Longint;
                                              Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  Text   : array[0..255] of Char;
  F      : TAlignment ;
  Grille : THgrid ;
  OK     : Boolean;
begin
  Grille := THgrid(Sender);
  StrPCopy(Text,Grille.Cells[Col, Row]);
  Grille.Canvas.Font.Style := Grille.Canvas.Font.Style - [fsItalic];
  OK := Grille.Cells[4, Row] = 'OK';

  if (gdFixed in State) then begin
    Grille.Canvas.Brush.Color := Grille.FixedColor ;
    Grille.Canvas.Font.Color  := Grille.Font.Color ;
    F:=taCenter
  end
  else begin
         if Col = 4 then F := taRightJustify
    else if Col = 5 then F := taCenter
                    else F := taLeftJustify;
    if (Col in [4, 5]) then begin
      if Row > 0 then begin
        if OK then Grille.Canvas.Font.Color  := clGreen
              else Grille.Canvas.Font.Color  := clRed;
      end;
    end
    else
      Grille.Canvas.Font.Color := clBlack;
    if (Row mod 2) = 0 then Grille.Canvas.Brush.Color := clWhite
                       else Grille.Canvas.Brush.Color := clSilver;
  end;

  case F of
    taRightJustify : ExtTextOut(Grille.Canvas.Handle, Rect.Right - Grille.Canvas.TextWidth(Grille.Cells[Col,Row])-3,
                               Rect.Top + 2,ETO_OPAQUE or ETO_CLIPPED, @Rect, Text, StrLen(Text), nil) ;
    taCenter       : ExtTextOut(Grille.Canvas.Handle,Rect.Left + ((Rect.Right-Rect.Left-Grille.canvas.TextWidth(Grille.Cells[Col,Row])) div 2),
                         Rect.Top + 2, ETO_OPAQUE or ETO_CLIPPED, @Rect, Text, StrLen(Text), nil) ;
    else
      ExtTextOut(Grille.Canvas.Handle, Rect.Left + 2, Rect.Top + 2, ETO_OPAQUE or ETO_CLIPPED,
                 @Rect, Text, StrLen(Text), nil) ;
  end;

  if ((gdfixed in State) and Grille.Ctl3D) then begin
    DrawEdge(Grille.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
    DrawEdge(Grille.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.AffDetail(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AffDetailEcr(TOB(FListe.Objects[0, Fliste.Row]));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.ChargerGrillePiece;
{---------------------------------------------------------------------------------------}
var  lTobPiece : TOB ;
     n         : integer;
     p         : integer;
     Lig       : integer;
begin

  // On commence par redessiner la grille
  DessinerGrille;

  FListe.RowCount := 2;
  Lig             := 0 ;

  // Parcours des pièces
  for n := 0 to TobGenerale.Detail.Count - 1 do begin

    lTobPiece := TobGenerale.Detail[ n ] ;

    // Parcours des lignes de la pièce
    for p := 0 to lTobPiece.Detail.Count - 1 do begin

      FListe.Cells[0, Lig + 1] := lTobPiece.Detail[p].GetValue('TE_NODOSSIER');
      FListe.Cells[1, Lig + 1] := lTobPiece.Detail[p].GetValue('E_JOURNAL');
      FListe.Cells[2, Lig + 1] := lTobPiece.Detail[p].GetValue('E_DATECOMPTABLE');
      FListe.Cells[3, Lig + 1] := lTobPiece.Detail[p].GetValue('E_GENERAL');
      FListe.Cells[4, Lig + 1] := lTobPiece.Detail[p].GetValue('E_REFINTERNE');
      FListe.Cells[5, Lig + 1] := StrFMontant( Arrondi( lTobPiece.Detail[p].GetValue('E_DEBIT'), V_PGI.OkDecV ), 15, V_PGI.OkDecV, '', True );
      FListe.Cells[6, Lig + 1] := StrFMontant( Arrondi( lTobPiece.Detail[p].GetValue('E_CREDIT'), V_PGI.OkDecV ), 15, V_PGI.OkDecV, '', True );

      FListe.Objects[0, Lig + 1] := lTobPiece.Detail[p];

      FListe.RowCount := FListe.RowCount + 1 ;
      Inc( Lig ) ;

    end

  end;

  FListe.RowCount := Lig + 1;


  {24/05/05 : Centrage des champs Debit et Crédit sur la droite}
  FListe.ColAligns[5] := taRightJustify;
  FListe.ColAligns[6] := taRightJustify;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLISTEECRCOMPTA.ChargerResultPiece;
{---------------------------------------------------------------------------------------}
var lTobPiece : Tob ;
    n : integer;
begin
  // On commence par redessiner la grille
  FListe.DefaultDrawing := False;
  DessinerGrille;

  // Remplissage de la grille
  FListe.RowCount := TobGenerale.Detail.Count + 1;
  for n := 0 to (TobGenerale.Detail.Count - 1) do begin

    lTobPiece := TobGenerale.Detail[ n ] ;

    FListe.Cells[0, n + 1] := lTobPiece.Detail[0].GetValue('TE_NODOSSIER');
    FListe.Cells[1, n + 1] := lTobPiece.Detail[0].GetValue('E_JOURNAL');
    FListe.Cells[2, n + 1] := lTobPiece.Detail[0].GetValue('E_GENERAL');
    FListe.Cells[3, n + 1] := lTobPiece.Detail[0].GetValue('E_REFINTERNE');
    FListe.Cells[4, n + 1] := lTobPiece.GetValue('RESULTAT');
    FListe.Cells[5, n + 1] := lTobPiece.GetValue('COMMENTAIRE');
  end;
  
end;

Initialization
  RegisterClasses ( [ TOF_TRLISTEECRCOMPTA ] ) ;
end.
