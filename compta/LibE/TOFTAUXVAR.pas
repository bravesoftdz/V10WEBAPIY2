{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : FTAUXVAR ()
Mots clefs ... : TOF;TOFTAUXVAR
*****************************************************************}
Unit TOFTAUXVAR;

Interface

Uses
     Controls,
     Classes,
     forms,
     Windows,
     sysutils,
     grids,
     HCtrls,
     HEnt1,
     graphics,
   {$IFNDEF EAGLCLIENT}
     FE_Main,
 {$ELSE}
     MaineAGL,
 {$ENDIF}
     Vierge,
     HTB97,
     UTOB,
     UTOF,
     UObjetEmprunt;

Type
  TOF_TAUXVAR = Class (TOF)
    private
       grTaux          : THGrid;
       fBoInSelectCell : Boolean;

       Procedure TauxVersGrille;
       Procedure GrilleVersTaux;
       Function  TauxValide(pInLigne : Integer) : Boolean;

       procedure grTauxGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
       procedure grTauxKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
       procedure grTauxSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
       procedure grTauxElipsisClick(Sender: TObject);

    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;

       procedure OnCancel                 ; Override ;
  end ;


const
   cColTVFix    = 0;
   cColTVdate   = 1;
   cColTVTaux   = 2;
   cColTVStatut = 3;

var
   gOmEmprunt : TObjetEmprunt;


Function SaisieTauxVariables(pObEmprunt : TObjetEmprunt) : Boolean;


Implementation


Function SaisieTauxVariables(pObEmprunt : TObjetEmprunt) : Boolean;
begin
  gOmEmprunt := pObEmprunt;
  Result := AGLLanceFiche('FP','FTAUXVAR','','','') <> '';
  gOmEmprunt := Nil;
end;



procedure TOF_TAUXVAR.OnArgument (S : String ) ;
begin
  Inherited;

  // Récup des controls
  grTaux := THGrid(GetControl('grTaux'));

  // Initialisation de la grille
  grTaux.DefaultRowHeight := 18;
  grTaux.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing,
                     goColSizing, goTabs, goThumbTracking, goTabs, goAlwaysShowEditor];
  grTaux.TwoColors := True;

  grTaux.FixedCols              := 1;
  grTaux.FixedRows              := 1;
  grTaux.ColCount               := 4;
  grTaux.RowCount               := 2;
  grTaux.Rowheights[0]          := 24;

  grTaux.ColWidths[cColTVFix]   := 24;
  grTaux.ColWidths[cColTVStatut]:= -1;
  grTaux.Colwidths[cColTVDate]  := Trunc((grTaux.Width - grTaux.ColWidths[cColTVFix] - 24) / 2);
  grTaux.Colwidths[cColTVTaux]  := grTaux.Colwidths[cColTVDate];

  grTaux.ColLengths[cColTVStatut]   := -1;
  grTaux.ColEditables[cColTVStatut] := False;

  grTaux.Cells[cColTVDate,0]    := 'Date';
  grTaux.Cells[cColTVTaux,0]    := 'Taux';

  grTaux.ColAligns[cColTVDate]  := taLeftJustify;
  grTaux.ColAligns[cColTVTaux]  := taRightJustify;

  grTaux.ColTypes[cColTVDate]   := 'D';
  grTaux.ColTypes[cColTVTaux]   := 'R';

  grTaux.ColFormats[cColTVDate] := ShortDateFormat;
  grTaux.ColFormats[cColTVTaux] := '##0.000';

  grTaux.GetCellCanvas  := grTauxGetCellCanvas;
  grTaux.OnKeyDown      := grTauxKeyDown;
  grTaux.OnSelectCell   := grTauxSelectCell;
  grTaux.OnElipsisClick := grTauxElipsisClick;

  TFVierge(Ecran).Retour := '';

  // remplissage de la grille
  TauxVersGrille;

end ;


// Nouvelle ligne
procedure TOF_TAUXVAR.OnNew ;
begin
  Inherited ;
  if not TauxValide(grTaux.Row) then Exit;
  grTaux.RowCount := grTaux.RowCount + 1;
  grTaux.Col := cColTVdate;
  grTaux.Row := grTaux.RowCount - 1;
end ;

// Suppression d'une ligne
procedure TOF_TAUXVAR.OnDelete ;
var
  lInRow : Integer;
begin
  Inherited ;

  if grTaux.RowCount <= 1 then Exit;
  lInRow := grTaux.Row;
  if grTaux.RowCount = 2 then
  begin
     grTaux.Cells[cColTVTaux,lInRow] := '';
     grTaux.Cells[cColTVDate,lInRow] := '';
     grTaux.Col := cColTVdate;
  end
  else
  begin
     if lInRow > 0 then
        grTaux.Row := lInRow - 1;
     if grTaux.RowCount > lInRow then
        grTaux.DeleteRow(lInRow);
  end;
end ;

// Enregistrer et fermer
procedure TOF_TAUXVAR.OnUpdate ;
begin
  Inherited ;
  GrilleVersTaux;
  TFVierge(Ecran).Retour := '*';
  Ecran.Close;
end ;

//
procedure TOF_TAUXVAR.OnLoad ;
begin
  Inherited ;
end ;

//  Annulation de la saisie
procedure TOF_TAUXVAR.OnCancel;
begin
  inherited;
  TauxVersGrille;
end;

//
procedure TOF_TAUXVAR.OnClose ;
begin
  Inherited ;
end ;


// Transfert les taux de l'objet emprunt vers la grille
Procedure TOF_TAUXVAR.TauxVersGrille;
var
   lInCpt  : Integer;
   lObTaux : Tob;
begin
   if gOmEmprunt.TauxVariables.Detail.Count = 0 then
      grTaux.RowCount := 2
   else
      grTaux.RowCount := gOmEmprunt.TauxVariables.Detail.Count + 1;

   for lInCpt := 0 to gOmEmprunt.TauxVariables.Detail.Count - 1 do
   begin
      lObTaux := gOmEmprunt.TauxVariables.Detail[lInCpt];
      grTaux.Cells[cColTVdate,lInCpt + 1] := DateToStr(lObTaux.GetValue('TXV_DATE'));
      grTaux.Cells[cColTVTaux,lInCpt + 1] := gOmEmprunt.FormaterMontant(lObTaux.GetValue('TXV_TAUX'));
   end;
end;


// Transfert les taux de la grille vers l'objet emprunt
Procedure TOF_TAUXVAR.GrilleVersTaux;
var
   lInCpt  : Integer;
   lObTaux : TOB;
begin
   gOmEmprunt.CreerTaux(0);
   for lInCpt := 1 to grTaux.RowCount - 1 do
      if TauxValide(lInCpt) then
      begin
         gOmEmprunt.CreerTaux(gOmEmprunt.TauxVariables.Detail.Count + 1);
         lObTaux := gOmEmprunt.TauxVariables.Detail[gOmEmprunt.TauxVariables.Detail.Count - 1];
         lObTaux.PutValue('TXV_CODEEMPRUNT', gOmEmprunt.GetValue('EMP_CODEEMPRUNT') );
         lObTaux.PutValue('TXV_DATE', StrToDate(grTaux.Cells[cColTVDate, lInCpt]) );
         lObTaux.PutValue('TXV_TAUX', Valeur(grTaux.Cells[cColTVTaux, lInCpt]) );
      end;

   gOmEmprunt.TauxVariables.Detail.Sort('TXV_DATE');
end;


// Retourne True si une ligne de taux est valide
Function TOF_TAUXVAR.TauxValide(pInLigne : Integer) : Boolean;
begin
   try
      StrToDate(grTaux.Cells[cColTVDate, pInLigne]);
      Valeur(grTaux.Cells[cColTVTaux, pInLigne]);
      Result := True;
   except
      Result := False;
   end;
end;


// Apparence de la grille
procedure TOF_TAUXVAR.grTauxGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
   if (ACol = 0) or
      (ARow = 0) then
   begin
      Canvas.Font.Name  := 'arial';
      Canvas.Font.Size  := 10;
      Canvas.Font.Style := [fsBold];
   end;
end;


// VK_DOWN en bas de la grille -> ajout d'une ligne
procedure TOF_TAUXVAR.grTauxKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
   if ((Key = VK_DOWN) or (Key = 34)) and
      (grTaux.Row = grTaux.RowCount - 1) then
      OnNew;

   if (grTaux.Col = cColTVDate) and (Key in [VK_LEFT,VK_RIGHT]) then
      grTaux.InvalidateCell(grTaux.Col,grTaux.Row);
end;

// Changement de cellule
procedure TOF_TAUXVAR.grTauxSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
begin
   if fBoInSelectCell then Exit;
   fBoInSelectCell := True;

   if (aRow <> grTaux.Row) and (Not TauxValide(grTaux.Row)) then
      OnDelete;
   grTaux.ElipsisButton := (ACol = cColTVDate);

   fBoInSelectCell := False;
end;


// Sélection d'une date
procedure TOF_TAUXVAR.grTauxElipsisClick(Sender: TObject);
begin
   V_PGI.ParamDateproc(grTaux);
end;




Initialization
  registerclasses ( [ TOF_TAUXVAR ] ) ;
end.

