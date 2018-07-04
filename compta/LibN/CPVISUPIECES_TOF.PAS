{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPVISUPIECES ()
Mots clefs ... : TOF;CPVISUPIECES
*****************************************************************}
Unit CPVISUPIECES_TOF ;

Interface

Uses
     Forms,
     StdCtrls,
     Controls,
     Windows,
     Classes,
     vierge,
     sysutils,
     graphics,
     grids,
{$IFDEF EAGLCLIENT}
     MaineAGL,          // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}  
     FE_Main,          // AGLLanceFiche
{$ENDIF EAGLCLIENT}
     uTob,
     ComCtrls,
     HCtrls,
     HTB97,
     AglInit,
     HEnt1,
     HMsgBox,
     HsysMenu,
     utilPGI,
     uLibECriture,
     uLibPieceCompta,
     UTOF ;

procedure CPLanceFiche_VisuPieces( vStArgs : string ) ;

Type
  TOF_CPVISUPIECES = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
  private

    FListe                             : THGrid ;

    FChamps                            : string ;
    FTobPieces                         : Tob ;
    FAction                            : TActionFiche ;

    // Données annexes
    FTobSoc                            : Tob ;
    FTobEtab                           : Tob ;

    procedure FListeDblClick      (Sender: TObject) ;
    procedure FListePostDrawCell  (ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure SetGridSep(ACol, ARow : Integer ;  Canvas : TCanvas ; bHaut : Boolean) ;

    procedure AnalyseParametre( vStArg : String ) ;
    procedure AffichePieces ;

    procedure ZoomLigne ( vRow : integer ) ;

    function  GetLibelleEtab ( vStCodeEtab, vStDossier : string ) : string ;
    function  GetLibelleSoc  ( vStCodeSoc : string ) : string ;

    procedure ParamFListe ;

  end ;

Implementation

uses      CPSaisiePiece_Tof ;


procedure CPLanceFiche_VisuPieces( vStArgs : string ) ;
begin
  AGLLanceFiche( 'CP', 'CPVISUPIECES', '', '', vStargs ) ;
end ;



procedure TOF_CPVISUPIECES.OnArgument (S : String ) ;
begin
  Inherited ;

  // Initialisation des composants
  FListe := THGrid( Getcontrol('FLISTE', True) ) ;

  // Ajout Evts
  FListe.OnDblClick       := FListeDblClick ;
  FListe.PostDrawCell     := FListePostDrawCell ;

  // initialisation des paramètres
  FAction    := taConsult ;
  FChamps    := 'E_SOCIETE;E_ETABLISSEMENT;E_JOURNAL;E_NUMEROPIECE;E_DATECOMPTABLE;E_REFINTERNE' ;
  FTobPieces := nil ;
  if Assigned( TFVierge(Ecran).laTof.laTob ) then
    FTobPieces := Tob( TFVierge(Ecran).laTof.laTob ) ;
  AnalyseParametre( S ) ;

  FTobSoc  := RecupInfosSocietes('SO_SOCIETE;SO_LIBELLE') ;
  FTobEtab := Tob.Create('$ETAB', nil, -1 );
  ParamFListe ;

  // Affichage des pièces:
  AffichePieces ;

  SetControlEnabled ('BVALIDER', True) ;
  SetControlProperty('BVALIDER', 'ModalResult', 0) ;

end ;

procedure TOF_CPVISUPIECES.OnClose ;
begin
  Inherited ;
  if Assigned( FTobSoc ) then
    FreeAndNil( FTobSoc ) ;
  if Assigned( FTobEtab ) then
    FreeAndNil( FTobEtab ) ;
end ;

procedure TOF_CPVISUPIECES.OnDisplay () ;
//var  HMT : TComponent ;
begin
  Inherited ;
{
  HMT := TForm( Ecran ).FindComponent('HMTrad');
  if (HMT<>nil) then
    THSystemMenu(HMT).ResizeGridColumns( FListe );
}
end ;


procedure TOF_CPVISUPIECES.AnalyseParametre(vStArg: String);
var lStChaine : string;
    lStVal    : string ;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt(vStArg);

    // Stocke le XXX de "MODE=XXX" dans la variable
    if (Pos('ACTION=', Uppercase(lStChaine)) > 0) then
      begin
      FAction   := StringToAction( lStChaine ) ;
      end

    // Stocke le XXX de "MODE=XXX" dans la variable
    else if (Pos('TITRE=', Uppercase(lStChaine)) > 0) then
      begin
      lStVal        := Copy( lStChaine, 7, length(lStChaine) ) ;
      Ecran.Caption := lStVal ;
      UpdateCaption( Ecran ) ;
      end

    // Paramètres de date de saisie
    else if (Pos('CHAMPS=', Uppercase(lStChaine)) > 0) then
      begin
      FChamps   := Copy( lStChaine, 8, length(lStChaine) ) ;
      end ;

  end;
  {JP 20/07/06 : Pour la boucle ci_dessus, les champs doivent être séparés par des ",".
                 On remplace donc les "," par des ";" pour la suite du traitement}
  FChamps := FindEtReplace(FChamps, ',', ';', True);
end;


procedure TOF_CPVISUPIECES.FListeDblClick(Sender: TObject);
begin
  ZoomLigne ( FListe.row ) ;
end;

procedure TOF_CPVISUPIECES.ZoomLigne(vRow: integer);
var lPieceCpt : TPieceCompta ;
    lInfoEcr  : TInfoEcriture ;
    lDossier  : string ;
    lStSoc    : string ;
    lTobEcr   : Tob ;
    lTobTmp   : Tob ;
begin

  // pas de pièces pour la ligne
  if vRow > FTobPieces.Detail.Count then Exit ;

  // Récup Tob + dossier
  lTobEcr  := FTobPieces.Detail[ vRow - 1 ];
  lDossier := V_PGI.SchemaName ;
  if EsTMultiSoc then
    if lTobEcr.GetNumChamp('SYSDOSSIER') > 0
      then lDossier := lTobEcr.GetString('SYSDOSSIER')
      else begin
           lStSoc  := lTobEcr.GetString('E_SOCIETE') ;
           lTobTmp := FTobSoc.FindFirst( ['SO_SOCIETE'], [ lStSoc ], False ) ;
           if Assigned( lTobTmp )
             then lDossier := lTobTmp.GetString('SCHEMANAME') ;
           end ;

  // Chargement pièce
  lInfoEcr  := TInfoEcriture.Create( lDossier ) ;
  lPieceCpt := TPieceCompta.CreerPiece( lInfoEcr ) ;
  // -- Champs obligatoires...
  lPieceCpt.PutEntete('E_JOURNAL',       lTobEcr.GetString('E_JOURNAL') ) ;
  lPieceCpt.PutEntete('E_EXERCICE',      lTobEcr.GetString('E_EXERCICE') ) ;
  lPieceCpt.PutEntete('E_DATECOMPTABLE', lTobEcr.GetDateTime('E_DATECOMPTABLE') ) ;
  lPieceCpt.PutEntete('E_NUMEROPIECE',   lTobEcr.GetInteger('E_NUMEROPIECE') ) ;
  lPieceCpt.PutEntete('E_QUALIFPIECE',   lTobEcr.GetString('E_QUALIFPIECE') ) ;
  // -- Chargement
  lPieceCpt.LoadFromSQL ;

  // Ouverture Saisie
  ModifiePieceCompta( lPieceCpt, FAction ) ;

end;

procedure TOF_CPVISUPIECES.AffichePieces;
var  HMT : TComponent ;
begin
  FTobPieces.PutGridDetail( FListe, False, False, FChamps, True ) ;

  HMT := TForm( Ecran ).FindComponent('HMTrad');
  if (HMT<>nil) then
    THSystemMenu(HMT).ResizeGridColumns( FListe );
  
end;


procedure TOF_CPVISUPIECES.SetGridSep(ACol, ARow: Integer; Canvas: TCanvas;  bHaut: Boolean);
var R : TRect ;
begin
Canvas.Brush.Color := clRed ;
Canvas.Brush.Style := bsSolid ;
Canvas.Pen.Color   := clRed ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psSolid ;
Canvas.Pen.Width   := 1 ;
R:=FListe.CellRect(ACol, ARow) ;
if bHaut then begin Canvas.MoveTo(R.Left, R.Top) ; Canvas.LineTo(R.Right+1, R.Top) end
         else begin Canvas.MoveTo(R.Left, R.Bottom-1) ; Canvas.LineTo(R.Right+1, R.Bottom-1) end ;
end;

procedure TOF_CPVISUPIECES.FListePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var lTobEcr : Tob ;
    lText       : string ;
    lRect       : TRect ;
begin

// initialisation  ??
//  Grille.Canvas.Font.Style := []; //Grille.Canvas.Font.Style-[fsItalic]-[fsBolt] ;
//  Grille.Canvas.Font.Color := clWindowText ;

  if FTobPieces.Detail.Count = 0 then Exit ;
  if Arow = 0 then Exit ;
  if Acol = 0 then Exit ;
  if Arow > FTobPieces.Detail.Count  then Exit ;

  lTobEcr := FTobPieces.Detail[ Arow-1 ] ;

 // En attente
  if lTobEcr.GetNumChamp('STYLE') > 0 then
    begin
    if lTobEcr.GetString('STYLE') = 'BOLD' then
      Canvas.Font.Style := [fsBold]
    else if lTobEcr.GetString('STYLE') = 'ITALIC' then
      Canvas.Font.Style := [fsItalic] ;
    end ;

  if lTobEcr.GetNumChamp('COLOR') > 0 then
    begin
    if lTobEcr.GetString('COLOR') = 'RED' then
      Canvas.Font.Color := clRed
    else if lTobEcr.GetString('COLOR') = 'BLUE' then
      Canvas.Font.Color := clBlue
    else if lTobEcr.GetString('COLOR') = 'GREEN' then
      Canvas.Font.Color := clGreen
    end ;

  lRect := FListe.CellRect(ACol,ARow) ;
  lText := FListe.Cells[ACol, ARow] ;
  if FListe.ColNames[ACol] = 'E_SOCIETE' then
    begin
    lText := GetLibelleSoc( lText ) ;
    end
  else if FListe.ColNames[ACol] = 'E_ETABLISSEMENT' then
    begin
    lText := GetLibelleEtab( lText, lTobEcr.GetString('SYSDOSSIER') ) ;
    end ;

  Canvas.FillRect( lRect );
  Canvas.TextRect( lRect, (lRect.Left+3), (lRect.Top+lRect.Bottom) div 2 -5 , lText );

  if EstMultiSoc and ( lTobEcr.GetNumChamp('SYSDOSSIER') > 0 ) then
    if ( ARow >= FTobPieces.Detail.Count ) or
       ( lTobEcr.GetString('SYSDOSSIER') <> FTobPieces.Detail[ Arow ].GetString('SYSDOSSIER') ) then
         SetGridSep( ACol, ARow, Canvas, False ) ;


end;

procedure TOF_CPVISUPIECES.OnUpdate;
begin
  inherited;
  ZoomLigne ( FListe.row ) ;
end;

function TOF_CPVISUPIECES.GetLibelleEtab( vStCodeEtab, vStDossier : string ) : string ;
var lQEtab   : TQuery ;
    lTobEtab : Tob ;
begin

  if vStDossier = V_PGI.schemaName then
    result := RechDom('TTETABLISSEMENT', vStCodeEtab, False )
  else
    begin

    lTobEtab := FTobEtab.FindFirst( ['SYSDOSSIER', 'ET_ETABLISSEMENT'], [ vStDossier, vStCodeEtab ], False ) ;
    if Assigned( lTobEtab ) then
      result := lTobEtab.GetString('ET_LIBELLE')
    else
      begin
      // Recherche libellé établissement
      lQEtab := OpenSelect( 'SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT = "' + vStCodeEtab + '"', vStDossier ) ;
      if not lQEtab.Eof
        then result := lQEtab.FindField('ET_LIBELLE').AsString
        else result := vStCodeEtab ;
      Ferme( lQEtab ) ;

      // Stockage du résultat
      lTobEtab := Tob.Create('ETABLISS', FTobEtab, -1 ) ;
      lTobEtab.AddChampSupValeur( 'SYSDOSSIER', vStDossier ) ;
      lTobEtab.PutValue('ET_ETABLISSEMENT',     vStCodeEtab ) ;
      lTobEtab.PutValue('ET_LIBELLE',           result ) ;

      end ;
    end ;

end;

function TOF_CPVISUPIECES.GetLibelleSoc(vStCodeSoc: string): string;
var lTobSoc : Tob ;
begin

  result := vStCodeSoc ;
  if not Assigned( FTobSoc ) then Exit ;
  {JP 30/05/07 : FQ 17963 : Ajout du code société en sus du libellé}
  lTobSoc := FTobSoc.FindFirst( ['SO_SOCIETE'], [ vStCodeSoc ], False ) ;
  if Assigned( lTobSoc )
   then result := vStCodeSoc + ' / ' + lTobSoc.GetString('SO_LIBELLE') ;

end;

procedure TOF_CPVISUPIECES.ParamFListe ;
var i       : integer ;
    lStChps : string ;
    lStCol  : string ;
    lNbCol  : integer ;
begin

  lstChps    := FChamps ;
  lNbCol     := 1 ;
  i          := 0 ;

  FListe.ColCount := 50 ;

  // Paramétrage largeur en fonction des champs...
//  for i := 1 to FListe.ColCount - 1 do
  while lstChps <> '' do
    begin
    i      := i + 1 ;
    lNbCol := lNbCol + 1 ;

    lStCol := ReadTokenSt( lStchps ) ;

    if lStCol = 'E_SOCIETE' then
      begin
      FListe.ColWidths[i] := 75 ;
      FListe.Cells[i, 0] := 'Société';
      end
    else if lStCol = 'E_ETABLISSEMENT' then
      begin
      FListe.ColWidths[i] := 75 ;
      FListe.Cells[i, 0] := 'Etab.';
{      if FListe.Cells[i, 1] = '' then // ???? pourquoi faire ????
        begin
        FListe.ColLengths[i] := -1;
        FListe.ColWidths[i] := -1;
        end}
      end
    else if lStCol = 'E_JOURNAL' then
      begin
      FListe.ColWidths[i] := 50 ;
      FListe.ColAligns[i] := taCenter;
      FListe.Cells[i, 0] := 'Journal';
      end
    else if lStCol = 'E_GENERAL' then
      begin
      FListe.ColWidths[i] := 100 ;
      FListe.Cells[i, 0] := 'Général';
      end
    else if lStCol = 'E_AUXILIAIRE' then
      begin
      FListe.ColWidths[i] := 100 ;
      FListe.Cells[i, 0] := 'Auxiliaire';
      end
    else if lStCol = 'E_REFINTERNE' then
      begin
      FListe.ColWidths[i] := 100 ;
      FListe.Cells[i, 0] := 'Référence';
      end
    else if lStCol = 'E_LIBELLE' then
      begin
      FListe.ColWidths[i] := 130 ;
      FListe.Cells[i, 0] := 'Libellé';
      end
    else if lStCol = 'E_DEBIT' then
      begin
      FListe.ColWidths[i]  := 50 ;
      FListe.ColTypes[i]   := 'D';
      FListe.ColFormats[i] := '#####0.00';
      FListe.ColAligns[i]  := taRightJustify;
      FListe.Cells[i, 0]   := 'Débit';
      end
    else if lStCol = 'E_CREDIT' then
      begin
      FListe.ColWidths[i]  := 50 ;
      FListe.ColTypes[i]   := 'D';
      FListe.ColFormats[i] := '#####0.00';
      FListe.ColAligns[i]  := taRightJustify;
      FListe.Cells[i, 0]   := 'Crédit';
      end
    else if lStCol = 'E_NUMEROPIECE' then
      begin
      FListe.ColWidths[i] := 50 ;
      FListe.ColTypes[i]  := 'I';
      FListe.ColFormats[i] := '#####0';
      FListe.ColAligns[i] := taRightJustify;
      FListe.Cells[i, 0] := 'Numéro';
      end
    else if lStCol = 'E_DATECOMPTABLE' then
      begin
      FListe.ColWidths[i]  := 80 ;
      FListe.ColAligns[i] := taCenter;
      FListe.ColTypes[i]   := 'D';
      FListe.ColFormats[i]  := 'dd/mm/yyyy';
      FListe.Cells[i, 0] := 'Date Cpt';
      end
    else if lStCol = 'SYSDOSSIER' then begin
      FListe.ColWidths[i] := 75;
      FListe.Cells[i, 0] := 'Dossier';
      if FListe.Cells[i, 1] = '' then begin
        FListe.ColLengths[i] := -1;
        FListe.ColWidths[i] := -1;
      end
    end

    else if lStCol = 'ERREUR' then begin
      FListe.ColWidths[i] := 243;
      FListe.Cells[i, 0] := 'Erreur';
      FListe.ScrollBars := ssBoth;
    end;

  end ;

  FListe.ColCount := lNbCol ;

end;

Initialization
  registerclasses ( [ TOF_CPVISUPIECES ] ) ;
end.
