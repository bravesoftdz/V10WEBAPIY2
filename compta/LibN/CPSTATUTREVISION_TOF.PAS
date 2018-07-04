{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSTATUTREVISION ()
Mots clefs ... : TOF;CPSTATUTREVISION
*****************************************************************}
Unit CPSTATUTREVISION_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
     PrintDBG, // PrintDBG
{$ENDIF}
     extctrls, // TImage
     Windows,  // TRect
     Grids,    // TGridDrawState
     Graphics, // TCanvas
     ed_Tools, // MoveCurProgressForm
     ParamSoc, // GetParamSocSecur
     Ent1,     // VH^.
     HSysMenu, // THSystemMenu
     Htb97,    // TToolBarbutton
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     uTob, 
     UTOF ;

Type

  TCPStatutRevision = class (TObject)

  private
    FTobGene        : Tob; // Liste des comptes généraux

    FLastError      : integer;
    FAvecAffichage  : Boolean;

    FStPlanRevision : string;
    FStCptBilanOuv  : string; // Compte de bilan d'ouverture
    FStCptBilanFer  : string; // Compte de bilan de fermeture

    procedure GenereErreur( vInErr : integer ; vTob : Tob );

  public
    FTobErreur : Tob;

    constructor Create;
    procedure   Execute;
    destructor  Destroy; override;

    property LastError     : integer read FLastError;
    property AvecAffichage : Boolean read FAvecAffichage  write FAvecAffichage;
    property PlanRevision  : string  read FStPlanRevision write FStPlanRevision;
  end;


  TOF_CPSTATUTREVISION = Class (TOF)

    FListe : THGrid;
    CBOk1  : TCheckBox;
    CBOk2  : TCheckBox;
    CBOk3  : TCheckBox;

    IM1    : TImage;
    IM2    : TImage;

    BCherche  : TToolBarButton97;
    BImprimer : TToolBarButton97;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;

    procedure OnDblClickFListe     ( Sender : TObject );
    procedure OnClickCB            ( Sender : TObject );
    procedure OnClickBCherche      ( Sender : TObject );
    procedure OnClickBImprimer     ( Sender : TObject );
    procedure OnKeyDownEcran       ( Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
    procedure OnKeyDownFListe      ( Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure OnPostDrawCellFListe (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState);

  private
    FTobFListe : Tob;
    FStPlanRevision : string;
    FStTitreEcran   : string;

    procedure CalculStatutRevision;
    procedure AfficheErreur;
    procedure StatutEcran( vBoARecalculer : boolean );
  public

  end ;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ControleStatutRevision;

Implementation

uses uTofConsEcr; // OperationsSurComptes

const cColErreur    = 1;
      cColCompte    = 2;
      cColLibelle   = 3;
      cColCycle     = 4;
      cColComptable = 5;
      cColMvt       = 6;
      cColSolde     = 7;

////////////////////////////////////////////////////////////////////////////////
{ TCPStatutRevision }
constructor TCPStatutRevision.Create;
begin
  FTobGene   := Tob.Create('GENE', nil, -1);
  FTobErreur := Tob.Create('ERR', nil, -1);

  FLastError      := 0;
  FAvecAffichage  := False;
  FStPlanRevision := '';

  FStCptBilanOuv := GetParamSocSecur('SO_OUVREBIL', '');
  FStCptBilanFer := GetParamSocSecur('SO_FERMEBIL', '');
end;

////////////////////////////////////////////////////////////////////////////////
destructor TCPStatutRevision.Destroy;
begin
  FreeAndNil(FTobErreur);
  FreeAndNil(FTobGene);
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TCPStatutRevision.Execute;
var i : integer;
    lTob : Tob;
    lDbTD : Double;
    lDbTC : Double;
begin
  if FStPlanRevision = '' then
    FStPlanRevision := VH^.Revision.Plan;

  if FStPlanRevision = '' then Exit;

  try
    FTobGene.ClearDetail;
    FTobGene.LoadDetailFromSQL('SELECT G_GENERAL, G_LIBELLE, G_TOTDEBE, G_TOTCREE, '+
                               'G_CYCLEREVISION FROM GENERAUX WHERE ' +
                               'G_GENERAL <> "' + FStCptBilanOuv + '" AND ' +
                               'G_GENERAL <> "' + FStCptBilanFer + '" ORDER BY G_GENERAL');

    InitMoveProgressForm(nil, 'Contrôle du statut du dossier (révision)', 'Traitement en cours...', FTobGene.Detail.Count-1, True, True);
    for i := 0 to FTobGene.Detail.Count -1 do
    begin
      lTob := FTobGene.Detail[i];
      MoveCurProgressForm('Compte général : ' +
                          lTob.GetString('G_GENERAL') + ' - ' +
                          lTob.GetString('G_LIBELLE'));

      if lTob.GetString('G_CYCLEREVISION') = '' then
      begin // Comptes non associés à un cycle de révision
        lDbTD := Arrondi(lTob.GetDouble('G_TOTDEBE'), 2);
        lDbTC := Arrondi(lTob.GetDouble('G_TOTCREE'), 2);

        if (lDbTD = 0) and (lDbTC = 0) then // non mouvementés
        begin
          GenereErreur( 1 , lTob );
        end
        else
          if ( lDbTD - lDbTC ) = 0 then // mouvementés mais soldés
          begin
            GenereErreur( 2 , lTob );
          end
          else
            if ( lDbTD - lDbTC ) <> 0 then // non soldés
            begin
              GenereErreur( 3 , lTob );
            end;
      end;
    end;

  finally
    FiniMoveProgressForm;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TCPStatutRevision.GenereErreur(vInErr: integer; vTob: Tob);
var lTobErreur : Tob;
    lDbTotal : Double;
begin
  lTobErreur := Tob.Create('', FTobErreur, -1);
  lTobErreur.AddChampSupValeur('ERREUR', vInErr);
  lTobErreur.AddChampSupValeur('GENERAL', vTob.GetString('G_GENERAL'));
  lTobErreur.AddChampSupValeur('LIBELLE', vTob.GetString('G_LIBELLE'));
  lTobErreur.AddChampSupValeur('CYCLE', vTob.GetString('G_CYCLEREVISION'));
  lDBTotal := Arrondi( vTob.GetDouble('G_TOTDEBE') - vTob.GetDouble('G_TOTCREE'), 2);
  if (lDbTotal) >= 0 then
    lTobErreur.AddChampSupValeur('COMPTABLE', StrFMontant(lDbTotal, 13, V_Pgi.OkDecV, '', True) + ' D')
  else
    lTobErreur.AddChampSupValeur('COMPTABLE', StrFMontant(Abs(lDbTotal), 13, V_Pgi.OkDecV, '', True) + ' C');

  case vInErr of

    1 : begin // non mouvementés
          lTobErreur.AddChampSupValeur('MVT', '-');
          lTobErreur.AddChampSupValeur('SOLDE', 'X');
        end;

    2 : begin // mouvementés, mais soldés
          lTobErreur.AddChampSupValeur('MVT', 'X');
          lTobErreur.AddChampSupValeur('SOLDE', 'X');
        end;

    3 : begin // mouvementés, et non soldés
          lTobErreur.AddChampSupValeur('MVT', 'X');
          lTobErreur.AddChampSupValeur('SOLDE', '-');
        end;
  end;

  if vInErr > FLastError then
    FLastError := vInErr;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ControleStatutRevision;
begin
  V_Pgi.ZoomOle := True;
  AGLLanceFiche('CP', 'CPSTATUTREVISION', '', '', '');
  V_Pgi.ZoomOle := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnArgument (S : String ) ;
begin
  inherited;

  Ecran.OnKeyDown := OnKeyDownEcran;

  FListe := THGrid(GetControl('FLISTE', True));
  FListe.PostDrawCell := OnPostDrawCellFListe;
  FListe.OnKeyDown  := OnKeyDownFListe;
  FListe.OnDblClick := OnDblClickFListe;

  FListe.ColAligns[cColErreur]    := TaCenter;
  FListe.ColAligns[cColCompte]    := TaCenter;
  FListe.ColAligns[cColLibelle]   := TaLeftJustify;
  FListe.ColAligns[cColCycle]     := TaCenter;
  FListe.ColAligns[cColComptable] := TaRightJustify;

  FListe.ColTypes[cColMvt]        := 'B';
  FListe.ColAligns[cColMvt]       := TaCenter;
  FListe.ColFormats[cColMvt]      := IntToStr( Byte( csCheckBox ) );

  FListe.ColTypes[cColSolde]      := 'B';
  FListe.ColAligns[cColSolde]     := TaCenter;
  FListe.ColFormats[cColSolde]    := IntToStr( Byte( csCheckBox ) );

  CBOk1 := TCheckBox(GetControl('CBOK1', True));
  CBOk2 := TCheckBox(GetControl('CBOK2', True));
  CBOk3 := TCheckBox(GetControl('CBOK3', True));

  IM1 := TImage(GetControl('IM1', True));
  IM2 := TImage(GetControl('IM2', True));

  CBOk1.OnClick := OnClickCB;
  CBOk2.OnClick := OnClickCB;
  CBOk3.OnClick := OnClickCB;

  BCherche  := TToolBarButton97(Getcontrol('BCHERCHE', True));
  BCherche.OnClick  := OnClickBCherche;

  BImprimer := TToolBarButton97(Getcontrol('BIMPRIMER', True));
  BImprimer.OnClick := OnClickBImprimer;

{$IFDEF EAGLCLIENT}
  BImprimer.Visible := False;
{$ENDIF}

  FTobFListe := Tob.Create('', nil, -1);
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... : 20/02/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnLoad ;
begin
  Inherited ;
  FStPlanRevision := VH^.Revision.Plan;
  Ecran.Caption := Ecran.Caption + ' ' + FStPlanRevision + ' - ' + RechDom('CPPLANREVISION', FStPlanRevision, False);
  FStTitreEcran := Ecran.Caption;
  UpdateCaption(Ecran);
  BCherche.Click;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnPostDrawCellFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TheRect : TRect;
begin
  if (ACol = cColCycle) or (ACol = cColComptable) then
  begin
    if ((Trim(FListe.Cells[ACol, ARow]) = '') or
        (Valeur(FListe.Cells[ACol, ARow]) = 0)) then
    begin
      TheRect := FListe.CellRect(ACol, ARow);
      Canvas.TextRect(TheRect, TheRect.Left, TheRect.Top, '');
      Canvas.Brush.Color := FListe.FixedColor;
      Canvas.Brush.Style := bsBDiagonal;
      Canvas.Pen.Color := FListe.FixedColor;
      Canvas.Pen.Mode := pmCopy;
      Canvas.Pen.Style := psClear;
      Canvas.Pen.Width := 1;
      Canvas.Rectangle(TheRect);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnDblClickFListe(Sender: TObject);
begin
  if FListe.Cells[cColCompte, FListe.Row] <> '' then
  begin
    OperationsSurComptes(FListe.Cells[cColCompte, FListe.Row], VH^.Encours.Code, '','', False);
    StatutEcran( True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of

    VK_F5 : if FListe.Focused then FListe.OnDblClick(Self);

    VK_F9 : BCherche.Click;

    80 :  if Shift = [ssCtrl] then // Ctrl + P
            BImprimer.Click;
  else
  end;
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if FListe.SortEnabled then
  begin
    if (Shift = [ssAlt]) and (Key in [Ord('1')..Ord('7')]) then
    begin
      case key of
        Ord('1') : FListe.SortGrid(1, (FListe.SortedCol = 1) xor FListe.SortDesc);
        Ord('2') : FListe.SortGrid(2, (FListe.SortedCol = 2) xor FListe.SortDesc);
        Ord('3') : FListe.SortGrid(3, (FListe.SortedCol = 3) xor FListe.SortDesc);
        Ord('4') : FListe.SortGrid(4, (FListe.SortedCol = 4) xor FListe.SortDesc);
        Ord('5') : FListe.SortGrid(5, (FListe.SortedCol = 5) xor FListe.SortDesc);
        Ord('6') : FListe.SortGrid(6, (FListe.SortedCol = 6) xor FListe.SortDesc);
        Ord('7') : FListe.SortGrid(7, (FListe.SortedCol = 7) xor FListe.SortDesc);
      else
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.CalculStatutRevision;
var lCPStatutRevision : TCPStatutRevision;
    lTobTemp : Tob;
    i : integer;
begin
  lCPStatutRevision := TCPStatutRevision.Create;
  try
    lCPStatutRevision.PlanRevision := FStPlanRevision;
    lCPStatutRevision.Execute;
    case lCPStatutRevision.FLastError of
      0,1 : VH^.BStatutRevision.ImageIndex := 0; // Vert
      2   : VH^.BStatutRevision.ImageIndex := 1; // Orange
      3   : VH^.BStatutRevision.ImageIndex := 2; // Rouge
      4   : VH^.BStatutRevision.ImageIndex := 3; // Bleu
    else
    end;
    // Mise à jour du paramsoc qui conserve la couleur du Statut
    SetParamSoc('SO_CPSTATUTREVISION', VH^.BStatutRevision.ImageIndex);
    
    FTobFListe.ClearDetail;
    if lCPStatutRevision.FTobErreur.Detail.Count > 0 then
    begin
      for i := 0 to lCPStatutRevision.FTobErreur.Detail.Count - 1 do
      begin
        lTobTemp := Tob.Create('', FTobFListe, -1);
        lTobTemp.Dupliquer(lCPStatutRevision.FTobErreur.Detail[i], True, True);
      end;

      AfficheErreur;
    end;
  finally
    lCPStatutRevision.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.AfficheErreur;
var i : integer;
    lBoOk1 : Boolean;
    lBoOk2 : Boolean;
    lBoOk3 : Boolean;
    lInErreur : integer;
begin
  if FTobFListe = nil then Exit;

  try
    lBoOk1 := CBOk1.Checked;
    lBoOk2 := CBOk2.Checked;
    lBoOk3 := CBOk3.Checked;

    FListe.VidePile(False);
    FListe.BeginUpdate;
    FListe.RowCount := 2;

    for i := 0 to FTobFListe.Detail.count -1 do
    begin
      lInErreur := FTobFListe.Detail[i].GetInteger('ERREUR');

      if (lBoOk1 and (lInErreur = 1)) or
         (lBoOk2 and (lInErreur = 2)) or
         (lBoOk3 and (lInErreur = 3)) then
      begin
        FListe.InsertRow(FListe.RowCount-1);
        FListe.Cells[cColErreur, FListe.RowCount-2]    := FTobFListe.Detail[i].Getstring('ERREUR');
        FListe.Cells[cColCompte, FListe.RowCount-2]    := FTobFListe.Detail[i].Getstring('GENERAL');
        FListe.Cells[cColLibelle, FListe.RowCount-2]   := FTobFListe.Detail[i].Getstring('LIBELLE');
        FListe.Cells[cColcycle, FListe.RowCount-2]     := FTobFListe.Detail[i].Getstring('CYCLE');
        FListe.Cells[cColComptable, FListe.RowCount-2] := FTobFListe.Detail[i].Getstring('COMPTABLE');
        FListe.Cells[cColMvt, FListe.RowCount-2]       := FTobFListe.Detail[i].Getstring('MVT');
        FListe.Cells[cColSolde, FListe.RowCount-2]     := FTobFListe.Detail[i].Getstring('SOLDE');
      end;
    end;

    if FListe.RowCount > 2 then
      FListe.RowCount := FListe.RowCount-1;

  finally
    FListe.Row := 1;
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
    FListe.EndUpdate ;
    if FListe.CanFocus then FListe.SetFocus;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnClickCB(Sender: TObject);
begin
  AfficheErreur;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnClickBCherche(Sender: TObject);
begin
  CalculStatutRevision;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSTATUTREVISION.OnClickBImprimer(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
  PrintDBGrid (FListe, nil, Ecran.Caption, '');
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/04/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSTATUTREVISION.StatutEcran(vBoARecalculer: boolean);
begin
  Ecran.Caption := FStTitreEcran;
  if vBoARecalculer then
  begin
    Ecran.Caption := Ecran.Caption + ' (à recalculer)';
    BCherche.Glyph := IM2.Picture.BitMap;
  end
  else
    BCherche.Glyph := IM1.Picture.BitMap;

  UpdateCaption(Ecran);
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPSTATUTREVISION ] ) ;
end.
