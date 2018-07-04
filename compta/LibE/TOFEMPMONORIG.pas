unit TOFEMPMONORIG;

Interface

Uses
     Controls,
     Graphics,
     Classes,
     Grids,
     forms,
     Windows,
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
  {$IFNDEF EAGLCLIENT}
     FE_Main,
 {$ELSE}
     MaineAGL,
 {$ENDIF}
     UTOF,
     TOFEMPIMP,
     uCreCommun,
     UObjetEmprunt;

Type
  TOF_EMPMONORIG = Class (TOF)
    private
       grEcheances       : THGrid;
       fOmEmpDev         : TObjetEmprunt;

       procedure grEcheancesGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
       Procedure bImprimerClick(Sender: TObject);

    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
  end ;

var
  gOmEmprunt : TObjetEmprunt;


Procedure EmpruntEnMonaieOrigine(pObEmprunt : TObjetEmprunt);



Implementation

Procedure EmpruntEnMonaieOrigine(pObEmprunt : TObjetEmprunt);
begin
   gOmEmprunt := pObEmprunt;
   AGLLanceFiche('FP','FEMPMONORIG','','','');
   gOmEmprunt := Nil;
end;


procedure TOF_EMPMONORIG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EMPMONORIG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_EMPMONORIG.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_EMPMONORIG.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EMPMONORIG.OnArgument (S : String ) ;
var
   lRdTaux         : Double;
begin
   inherited;

   Ecran.Caption := Ecran.Caption + ' (' + gLibDevSaisie + ')';

   // Récup et init des controles
   if Assigned(GetControl('BIMPRIMER1')) then
      TToolBarButton97(GetControl('BIMPRIMER1')).OnClick := bImprimerClick;

   SetControlProperty('EMP_TOTINT',    'CurrencySymbol', gSymbDevSaisie);
   SetControlProperty('EMP_TOTASS',    'CurrencySymbol', gSymbDevSaisie);
   SetControlProperty('EMP_TOTAMORT',  'CurrencySymbol', gSymbDevSaisie);
   SetControlProperty('EMP_TOTINTASS', 'CurrencySymbol', gSymbDevSaisie);
   SetControlProperty('EMP_TOTVERS',   'CurrencySymbol', gSymbDevSaisie);

   // Init de la grille des échéances
   grEcheances                  := THGrid(GetControl('GRECHEANCES'));
   grEcheances.GetCellCanvas    := grEcheancesGetCellCanvas;
   grEcheances.DefaultRowHeight := 18;
   grEcheances.ColCount         := cNbColEcheances;
   grEcheances.RowCount         := 2;
   grEcheances.FixedCols        := 1;
   grEcheances.FixedRows        := 1;
   grEcheances.Options          := [goRowSelect, goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goThumbTracking];
   grEcheances.RowHeights[0]    := 22;

   grEcheances.ColWidths[cColNumEch]        := 36;
   grEcheances.ColWidths[cColDate]          := cInLargColDate;
   grEcheances.ColWidths[cColTxInteret]     := cInLargColTaux;
   grEcheances.ColWidths[cColAmortissement] := cInLargColMnt;
   grEcheances.ColWidths[cColInteret]       := cInLargColMnt;
   grEcheances.ColWidths[cColAssurance]     := cInLargColMnt;
   grEcheances.ColWidths[cColTotInteret]    := cInLargColMnt;
   grEcheances.ColWidths[cColVersement]     := cInLargColMnt;
   grEcheances.ColWidths[cColSolde]         := cInLargColMnt;
   grEcheances.ColWidths[cColReport]        := -1;
   grEcheances.ColWidths[cColDette]         := -1;
   grEcheances.ColWidths[cColCumulA]        := -1;
   grEcheances.ColWidths[cColCumulI]        := -1;
   grEcheances.ColWidths[cColCumulM]        := -1;
   grEcheances.ColWidths[cColTypeLigne]     := -1;
   grEcheances.ColWidths[cColModif]         := -1;

   if gOmEmprunt.GetValue('EMP_EMPTYPEASSUR') = taTaux then
   begin
      grEcheances.ColWidths[cColTxAssurance]  := cInLargColTaux;
      grEcheances.ColWidths[cColTxGlobal]     := cInLargColTaux;
   end
   else
   begin
      grEcheances.ColWidths[cColTxAssurance]  := -1;
      grEcheances.ColWidths[cColTxGlobal]     := -1;
   end;

   grEcheances.ColAligns[cColNumEch]        := taCenter;
   grEcheances.ColAligns[cColDate]          := taRightJustify;
   grEcheances.ColAligns[cColTxInteret]     := taRightJustify;
   grEcheances.ColAligns[cColTxAssurance]   := taRightJustify;
   grEcheances.ColAligns[cColTxGlobal]      := taRightJustify;
   grEcheances.ColAligns[cColAmortissement] := taRightJustify;
   grEcheances.ColAligns[cColInteret]       := taRightJustify;
   grEcheances.ColAligns[cColAssurance]     := taRightJustify;
   grEcheances.ColAligns[cColTotInteret]    := taRightJustify;
   grEcheances.ColAligns[cColVersement]     := taRightJustify;
   grEcheances.ColAligns[cColSolde]         := taRightJustify;
   grEcheances.ColAligns[cColReport]        := taRightJustify;
   grEcheances.ColAligns[cColDette]         := taRightJustify;
   grEcheances.ColAligns[cColCumulA]        := taRightJustify;
   grEcheances.ColAligns[cColCumulI]        := taRightJustify;
   grEcheances.ColAligns[cColCumulM]        := taRightJustify;

   // Création de l'emprunt en monnaie d'origine
   fOmEmpDev := TObjetEmprunt.CreerEmprunt;
   fOmEmpDev.Assigner(gOmEmprunt);
   fOmEmpDev.ValeurPK := - gOmEmprunt.GetValue('EMP_CODEEMPRUNT');
   fOmEmpDev.PutValue('EMP_CODEEMPDEV', gOmEmprunt.GetValue('EMP_CODEEMPRUNT') );
   fOmEmpDev.PutValue('EMP_DEVISE', gCodeDevSaisie );
   fOmEmpDev.PutValue('EMP_DEVISESAISIE', '@' );

   // Conversion des montants dans la devise d'origine
   if gCodeDevSaisie = gCodeDevEuro then
      lRdTaux := 1/V_PGI.TauxEuro
   else
      lRdTaux := V_PGI.TauxEuro;
   fOmEmpDev.PutValue('EMP_CAPITAL', fOmEmpDev.GetValue('EMP_CAPITAL') * lRdTaux );
   if fOmEmpDev.GetValue('EMP_EMPTYPEASSUR') = taMontant then
      fOmEmpDev.PutValue('EMP_VALASSUR', fOmEmpDev.GetValue('EMP_VALASSUR') * lRdTaux );

   // simulation
   if fOmEmpDev.GetValue('EMP_EMPTYPEVERS') = tvVariable then
      fOmEmpDev.PutValue('EMP_AMORTCST', fOmEmpDev.CalcAmortCst);

   fOmEmpDev.Simulation(trVersement);
   fOmEmpDev.EcheancesVersGrille(grEcheances);

   // Totaux de l'emprunt
   SetControlProperty('EMP_TOTINT',    'Value', fOmEmpDev.GetValue('EMP_TOTINT') );
   SetControlProperty('EMP_TOTASS',    'Value', fOmEmpDev.GetValue('EMP_TOTASS') );
   SetControlProperty('EMP_TOTAMORT',  'Value', fOmEmpDev.GetValue('EMP_TOTAMORT') );
   SetControlProperty('EMP_TOTVERS',   'Value', fOmEmpDev.GetValue('EMP_TOTVERS') );
   SetControlProperty('EMP_TOTINTASS', 'Value', fOmEmpDev.GetValue('EMP_TOTINT') + fOmEmpDev.GetValue('EMP_TOTASS') );
{
      // taille de certains labels
      for lInCpt := 1 to 5 do
      begin
         laLabel := THLabel(GetControl('LALABEL'+inttostr(lInCpt)));
         if Assigned(laLabel) then
         begin
            lInRight := laLabel.Left + laLabel.Width;
            laLabel.AutoSize := True;
            laLabel.Left := lInRight - laLabel.Width;
         end;
      end;
}
end;

procedure TOF_EMPMONORIG.OnClose ;
begin
  fOmEmpDev.Free;
  Inherited ;
end ;



// ----------------------------------------------------------------------------
// Nom    : grEcheancesGetCellCanvas
// Date   : 26/03/2002
// Auteur : D. ZEDIAR
// Objet  : Style de la cellule
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_EMPMONORIG.grEcheancesGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

   // Cellules fixes en gras (titres)
   if (ACol = 0) or
      (ARow = 0) then
   begin
      Canvas.Font.Name  := 'Arial';
      Canvas.Font.Size  := 10;
      Canvas.Font.Style := [fsBold];
   end
   else
      if grEcheances.Cells[cColTypeLigne,ARow] = cStLigneTotal then
      begin
         // Totaux d'exercice
         Canvas.Brush.Color := cInFondExColor;
         Canvas.Font.Style  := [fsBold];
      end;

end;



Procedure TOF_EMPMONORIG.bImprimerClick(Sender: TObject);
begin
   if not EstSerialise then Exit;

   if fOmEmpDev.Echeances.Detail.Count = 0 then
   begin
      PGIBox('Le tableau d''amortissement est vide.',TitreHalley);
      Exit;
   end;

   try
      // Enregistrement de l'emprunt en devise d'origine
      if transactions(fOmEmpDev.DBSupprimeEtRemplaceEmprunt,1) <> oeOk then
         Exit;

      // Impression
      ImprimeEmprunt(fOmEmpDev.GetValue('EMP_CODEEMPRUNT'),fOmEmpDev.GetValue('EMP_LIBEMPRUNT')); {YMO 18/04/07 Ajout libellé}

   finally
      // suppression de l'emprunt en devise d'origine si nécéssaire
      transactions(fOmEmpDev.DBSupprimerEmprunt,1);
   end;
end;


Initialization
  registerclasses ( [ TOF_EMPMONORIG ] ) ;
end.

