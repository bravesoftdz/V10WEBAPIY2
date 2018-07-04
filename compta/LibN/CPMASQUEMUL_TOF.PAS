{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMASQUEMUL ()
Mots clefs ... : TOF;CPMASQUEMUL
*****************************************************************}
Unit CPMASQUEMUL_TOF ;

Interface

Uses
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     Graphics, // TCanvas
     Grids,
{$ELSE}
     dbGrids,
     Grids,
     hdb,
{$ENDIF}
     HCtrls,     // ReadTokenSt, THGrid
     windows,    // TRect
     HQry,       // THQuery
     ComCtrls,   // TPageControl
     UTOF ;

Type
  TOF_CPMASQUEMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    // Composants
    Q : THQuery ;

    {$IFDEF EAGLCLIENT}
      FListe: THGrid;
    {$ELSE}
      FListe: THDBGrid;
    {$ENDIF}

    Pages : TPageControl;

    // Variables
    FStAction : string; // Stocke XXX de l'argument "ACTION=XXX"
    FStType   : string; // Stocke le type de masque concerné

    // Gestion des Evts
    {$IFDEF EAGLCLIENT}
//      MulDrawCell       : TDrawCellEvent ;
    {$ELSE}
      MULDrawColumnCell : TDrawColumnCellEvent ;
    {$ENDIF}

    // Evènement de la fiche
    procedure FListeDblClick   ( Sender : TObject );
    procedure BInsertClick     ( Sender : TObject );
    procedure BSupprimerClick  ( Sender : TObject );
    procedure BChercheClick    ( Sender : TObject );

    // affichage de la grille pour renommé le titre des colonnes paramétrables
    {$IFDEF EAGLCLIENT}
//      procedure FListePostDrawCell (ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
//      procedure FListeOnDrawCell   (Sender: TObject; ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);
    {$ELSE}
      procedure FListeDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    {$ENDIF}


    // Initialisation de la fiche
    procedure AnalyseParametre(vStArg: string);
    procedure InitComposants;
    procedure InitEvenements;
    procedure InitEcran;

  {$IFDEF EAGLCLIENT}
    procedure MAJTitreCol ;
  {$ENDIF EAGLCLIENT}
  end;


function CPLanceFiche_MasqueMul( vStRange, vStLequel, vStArgs : string ) : string;

Implementation

uses
  {$IFDEF VER150} Variants, {$ENDIF}    // VarToStr
  {$IFDEF EAGLCLIENT}
       MainEagl,                        // AglLanceFiche
  {$ELSE}
       mul,
       FE_Main,                         // AglLanceFiche
  {$ENDIF}
  sysutils,                             // Uppercase
  paramSoc,                             // GetParamSocSecur
  HMsgBox,                              // PGIInfo
  CPMASQUESAISIE_TOF,                   // CInitMasqueDEChamps
  HTB97,                                // TToolBarButton97
  Controls;                             // mrYes


function CPLanceFiche_MasqueMul( vStRange, vStLequel, vStArgs : string ) : string;
begin
  result := AGLLanceFiche('CP', 'CPMASQUEMUL', vStRange, vStLequel, vStArgs ) ;
end ;


procedure TOF_CPMASQUEMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUEMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUEMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUEMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUEMUL.OnArgument (S : String ) ;
begin
  Inherited ;

  CInitMasqueDEChamps ;

  AnalyseParametre(S);
  InitComposants;
  InitEvenements;
  InitEcran;

end ;

procedure TOF_CPMASQUEMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUEMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUEMUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUEMUL.AnalyseParametre(vStArg: string);
var
  lStChaine: string;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt(vStArg);

    // Stocke le XXX de "ACTION=XXX" dans la variable Action
    if Pos('ACTION=', Uppercase(lStChaine)) > 0 then
      FStAction := Copy(lStChaine, 8, length(lStChaine))

    // Type de masque concerné
    else if Pos('CMS_TYPE=', Uppercase(lStChaine)) > 0 then
      FStType := Copy(lStChaine, 10, length(lStChaine))

    ;

    end;

  if FStType='' then
    FStType := 'SAI' ;

end;

procedure TOF_CPMASQUEMUL.BInsertClick(Sender: TObject);
begin
  CPLanceFiche_MasqueSaisie( '', '', 'ACTION=CREATION;CMS_TYPE=SAI') ;
  BChercheClick(nil) ;
end;

procedure TOF_CPMASQUEMUL.BSupprimerClick(Sender: TObject);
begin

  if FStAction='CONSULTATION' then Exit ;

  // Le résultat de la recherche est-il vide ?
  {$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then
    {$ELSE}
  if (Q.Eof and Q.Bof) then
    {$ENDIF}
  begin
    PGIInfo('Le résultat de la recherche est vide. Veuillez relancer une recherche.', Ecran.Caption);
    exit;
  end;

  // Confirmation
  if PGIAsk('Confirmez-vous la suppression du masque sélectionné (' + GetField('CMS_LIBELLE') + ') ?',
            Ecran.Caption) <> mrYes then Exit;

  // Supppression
  if executeSql('DELETE FROM CMASQUESAISIE WHERE CMS_TYPE="SAI" AND CMS_NUMERO=' + VarToStr(GetField('CMS_NUMERO')) ) <> 1
    then PGIInfo('Le masque de saisie sélectionné n''a pas pu être supprimé.', Ecran.Caption)
    else executeSql('DELETE FROM CMASQUECRITERES WHERE CMC_TYPE="SAI" AND CMC_NUMERO=' + VarToStr(GetField('CMS_NUMERO')) ) ;

  // On relnce la recherche
  TFMul(Ecran).BChercheClick(nil);

end;

procedure TOF_CPMASQUEMUL.FListeDblClick(Sender: TObject);
begin

  {$IFDEF EAGLCLIENT}
    if (Q.TQ.Detail.Count = 0) then Exit ;
  {$ELSE}
    if (Q.Eof and Q.Bof) then Exit ;
  {$ENDIF}

  CPLanceFiche_MasqueSaisie('', VarToStr(GetField('CMS_NUMERO')), 'ACTION=' + FStAction);
  BChercheClick(nil) ;

end;

procedure TOF_CPMASQUEMUL.InitComposants;
begin
  Q := THQuery(GetControl('Q', True));
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe', True));
  {$ELSE}
  FListe := THDBGrid(GetControl('FListe', True));
  {$ENDIF}
  Pages := TPageControl(GetControl('Pages', True));

end;

procedure TOF_CPMASQUEMUL.InitEcran;
begin
  // Cacher l'onglet "compléments"
  TTabSheet(GetControl('PCOMPLEMENT', True)).TabVisible := False;

  // Faire apparaître les boutons de création / suppression si besoin
  SetControlVisible('BINSERT',    (FStAction <> 'CONSULTATION') ) ;
  SetControlVisible('BSUPPRIMER', (FStAction <> 'CONSULTATION') ) ;

  // Init des zones de critères paramétrables
  CInitCritMasque( self ) ;
  {$IFDEF EAGLCLIENT}
  MAJTitreCol ;
  {$ENDIF EAGLCLIENT}


  // Filtre sur le type de saisie
  SetControlText('CMS_TYPE', FStType ) ;

end;

procedure TOF_CPMASQUEMUL.InitEvenements;
begin
  // Double click sur la liste
  FListe.OnDblClick := FListeDblClick;
  {$IFDEF EAGLCLIENT}
//    FListe.PostDrawCell     := FListePostDrawCell ;
//    MulDrawCell             := FListe.OnDrawCell ;
//    FListe.OnDrawCell       := FListeOnDrawCell;
  {$ELSE}
    MULDrawColumnCell       := FListe.OnDrawColumnCell ;
    FListe.OnDrawColumnCell := FListeDrawColumnCell ;
  {$ENDIF EAGLCLIENT}

  // Bouton de création
  TToolBarButton97(GetControl('BINSERT', True)).OnClick    := BInsertClick;
  TToolBarButton97(GetControl('BSUPPRIMER', True)).OnClick := BSupprimerClick;
  TToolBarButton97(GetControl('BCHERCHE', True)).OnClick := BChercheClick;

end;

{$IFDEF EAGLCLIENT}
{
procedure TOF_CPMASQUEMUL.FListePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var lText  : string ;
    lRect  : TRect ;
    lStCol : string;
    lStChp : string ;
    lIdx   : string ;
begin

  // MAJ du TITRE
  if ARow = 0 then
    begin

    lStCol   := FListe.ColNames[ ACol ] ;
    if pos( 'CMS_CRITERE', lstcol ) > 0 then
      begin
      lText := GetControlText('T' + lstcol ) ;
      lRect := FListe.CellRect(ACol,ARow) ;

      if lText <> '' then
        begin
        Canvas.FillRect( lRect );
        Canvas.TextRect( lRect,
                         lRect.Left  + ((lRect.Right - lRect.Left - canvas.TextWidth( lText ) ) div 2 ),
                         (lRect.Top+lRect.Bottom) div 2 -5 ,
                         lText );
        end ;
      end ;
    end
  else
  // MAJ des libellés des critères
    begin
    lStChp := FListe.ColNames[ACol] ;
    if pos('CMC_CRITERE', lStChp ) > 0 then
      begin
      lIdx := Copy( lStChp, length( lStChp ), 1 ) ;


      if GetParamSocSecur('SO_CPMASQUECRIT' + lIdx, '000') <> '000' then
        begin
        lText := FListe.Cells[ACol, ARow ] ;
        if lText <> ''
          then  lText := RechDom( ThMultiValComboBox(GetControl( lStChp , True)).DataType, lText, False )
          else  lText := '<<Tous>>' ;
        if lText <> '' then
          begin
          lRect := FListe.CellRect(ACol,ARow) ; ;
          Canvas.FillRect( lRect );
          Canvas.TextRect( lRect,
                           lRect.Left  + ((lRect.Right - lRect.Left - canvas.TextWidth( lText ) ) div 2 ),
                           (lRect.Top+lRect.Bottom) div 2 -5 ,
                           lText );
          end ;
        end ;
      end ;
    end ;

end;


procedure TOF_CPMASQUEMUL.FListeOnDrawCell(Sender: TObject; ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState);
var lText  : string ;
    lRect  : TRect ;
    lStCol : string;
    lStChp : string ;
    lIdx   : string ;
    lCanvas : TCanvas ;
begin

  if Assigned( MULDrawCell ) then
    MULDrawCell(Sender, ACol, ARow, ARect, AState) ;

  // MAJ du TITRE
  if ARow = 0 then
    begin
    lStCol   := FListe.ColNames[ ACol ] ;
    if pos( 'CMS_CRITERE', lstcol ) > 0 then
      begin
      lText   := GetControlText('T' + lstcol ) ;
      lRect   := ARect ;
      lCanvas := FListe.Canvas ;
      if lText <> '' then
        begin
        lCanvas.FillRect( lRect );
        lCanvas.TextRect( lRect,
                         lRect.Left  + ((lRect.Right - lRect.Left - lCanvas.TextWidth( lText ) ) div 2 ),
                         (lRect.Top+lRect.Bottom) div 2 -5 ,
                         lText );
        end ;
      end ;
    end
  else
  // MAJ des libellés des critères
    begin
    lStChp := FListe.ColNames[ACol] ;
    if pos('CMC_CRITERE', lStChp ) > 0 then
      begin
      lIdx := Copy( lStChp, length( lStChp ), 1 ) ;
      if GetParamSocSecur('SO_CPMASQUECRIT' + lIdx, '000') <> '000' then
        begin
        lText := FListe.Cells[ACol, ARow ] ;
        if lText <> ''
          then  lText := RechDom( ThMultiValComboBox(GetControl( lStChp , True)).DataType, lText, False )
          else  lText := '<<Tous>>' ;
        if lText <> '' then
          begin
          lRect   := ARect ;
          lCanvas := FListe.Canvas ;
          lCanvas.FillRect( lRect );
          lCanvas.TextRect( lRect,
                           lRect.Left  + ((lRect.Right - lRect.Left - lCanvas.TextWidth( lText ) ) div 2 ),
                           (lRect.Top+lRect.Bottom) div 2 -5 ,
                           lText );
          end ;
        end ;
      end ;
    end ;
end;
}

{$ELSE}
procedure TOF_CPMASQUEMUL.FListeDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer;
                                                Column: TColumn; State: TGridDrawState ) ;
var lIdx  : string ;
    lTxt  : string ;
    lRect : TRect ;
begin

  if Assigned( MULDrawColumnCell ) then
    MULDrawColumnCell(Sender, Rect, DataCol, Column, State) ;

  if pos('CMC_CRITERE',Column.FieldName)>0 then
    begin
    // Titre
    Column.Title.Caption := GetControlText('T' + Column.FieldName ) ;

    // Valeur
    if GetDataSet.Eof then Exit ;
    lIdx := Copy( Column.FieldName, length( Column.FieldName ), 1 ) ;
    lTxt := FListe.Columns[DataCol].Field.AsString ;
    if GetParamSocSecur('SO_CPMASQUECRIT' + lIdx, '000') <> '000' then
      begin
      if lTxt <> ''
        then  lTxt := RechDom( ThMultiValComboBox(GetControl(Column.FieldName, True)).DataType, lTxt, False )
        else  lTxt := '<<Tous>>' ;
      if (lTxt <> '') and (lTxt <> 'Error') then
        begin
        lRect := Rect ;
        FListe.Canvas.FillRect( lRect );
        FListe.Canvas.TextRect( lRect, (lRect.Left+3), (lRect.Top+lRect.Bottom) div 2 -5 , lTxt );
        end ;
      end ;
    end ;
{
  else if pos('CMS_CRITERE2',Column.FieldName)>0 then
    begin
    Column.Title.Caption := GetControlText('T' + Column.FieldName ) ;
    end

  else if pos('CMS_CRITERE3',Column.FieldName)>0 then
    begin
    Column.Title.Caption := GetControlText('T' + Column.FieldName ) ;
    end ;
}
end;
{$ENDIF}

procedure TOF_CPMASQUEMUL.BChercheClick(Sender: TObject);
begin

//  FListe.OnDrawCell       := FListeOnDrawCell;
  TFMul(Ecran).BChercheClick(Sender);

  {$IFDEF EAGLCLIENT}
  MAJTitreCol ;
  {$ENDIF EAGLCLIENT}
end;


{$IFDEF EAGLCLIENT}
procedure TOF_CPMASQUEMUL.MAJTitreCol;
var lCol   : integer ;
    lRow   : integer ;
    lStChp : string ;
    lIdx   : string ;
    lText  : string ;
begin

  for lCol := 1 to FListe.ColCount - 1 do
    begin
    lStChp   := FListe.ColNames[ lCol ] ;
    if pos( 'CMC_CRITERE', lStChp ) > 0 then
      begin
      for lRow := 0 to FListe.RowCount - 1 do
        begin
        // Titre
        if lRow = 0 then
          begin
          if GetControlVisible( 'T' + lStChp ) then
            FListe.Cells[ lCol , 0 ] := GetControlText('T' + lStChp ) ;
          end
        // Libellés
        else
          begin
          lIdx := Copy( lStChp, length( lStChp ), 1 ) ;
          if GetParamSocSecur('SO_CPMASQUECRIT' + lIdx, '000') <> '000' then
            begin
            lText := FListe.Cells[ lCol, lRow ] ;
            if lText <> ''
             then  lText := RechDom( ThMultiValComboBox(GetControl( lStChp , True)).DataType, lText, False )
             else  lText := '<<Tous>>' ;
            if lText <> '' then
              FListe.Cells[ lCol , lRow ] := lText ;
            end ;
          end
        end ;
      end ;
    end ;

//  FListe.Invalidate ;

end;
{$ENDIF EAGLCLIENT}

Initialization
  registerclasses ( [ TOF_CPMASQUEMUL ] ) ;
end.
