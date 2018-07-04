{***********UNITE*************************************************
Auteur  ...... : Laroche Vincent
Créé le ...... : 30/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULSUPPRANA ()
Mots clefs ... : TOF;CPMULSUPPRANA
*****************************************************************}
Unit CPMULSUPPRANA_TOF ;

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
     HDB,  // THDBGrid
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97,
     HPanel,     // THPanel
     HStatus,    // MoveCur, InitMove, FiniMove
     UtilPGI,    // _BlocageMonoPoste
     Ent1,
     DelVisuE,   // VisuPiecesGenere
     SaisUtil,   // EcrAna
     ED_TOOLS,   // VideListe
     SaisODA,    // TrouveEtLanceSaisieODA
     SaisComm,   // TTypeExo
     uTob,
     UtilSais    // MajSoldeSectionTOB
     ;

procedure CPLanceFicheMulSupprAna; // DetruitAnalytiques

Type
  TOF_CPMULSUPPRANA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Y_AXE, Y_EXERCICE, Y_JOURNAL : THValComboBox;
    Y_SECTION, Y_DATECOMPTABLE, Y_DATECOMPTABLE_ : THedit;
    ListePieces : TList;
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}
    procedure Y_EXERCICEChange(Sender: TObject);
    procedure Y_AXEChange(Sender: TObject);
    procedure BDetruitesClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);

    function  DetruitSelection : boolean;
    procedure DetruitLaPiece;
    procedure InitCriteres ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  uLibEcriture ; // WhereEcritureTOB

procedure CPLanceFicheMulSupprAna; // DetruitAnalytiques
begin
  if Not _BlocageMonoPoste(False) then Exit ;
  AGLLanceFiche('CP','CPMULSUPPRANA','','','');
end;

procedure TOF_CPMULSUPPRANA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMULSUPPRANA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMULSUPPRANA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPMULSUPPRANA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPMULSUPPRANA.OnArgument (S : String ) ;
begin
  Inherited ;

  ListePieces := TList.Create;
  PositionneEtabUser(GetControl('Y_ETABLISSEMENT',True));

  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE', True));
  {$ELSE}
  FListe := THDBGrid(GetControl('FLISTE', True));
  {$ENDIF}
  FListe.OnDblClick := FListeDblClick;

  Y_JOURNAL := THValComboBox(GetControl('Y_JOURNAL', True));
  Y_EXERCICE := THValComboBox(GetControl('Y_EXERCICE', True));
  Y_EXERCICE.OnChange := Y_EXERCICEChange;
  Y_DATECOMPTABLE  := THEdit(GetControl('Y_DATECOMPTABLE', True));
  Y_DATECOMPTABLE_ := THEdit(GetControl('Y_DATECOMPTABLE_', True));
  Y_AXE := THValComboBox(GetControl('Y_AXE', True));
  Y_AXE.OnChange := Y_AXEChange;
  Y_SECTION := THEdit(GetControl('Y_SECTION', True));
  TToolBarButton97(GetControl('BDETRUITES', True)).OnClick := BDetruitesClick;
  TToolBarButton97(GetControl('BZOOMPIECE', True)).OnClick := FListeDblClick;
  TFMul(Ecran).BOuvrir.OnClick := BOuvrirClick;
  Initcriteres;
end ;

procedure TOF_CPMULSUPPRANA.OnClose ;
begin
  Inherited ;

  if ListePieces.Count>0 then begin
    if PGIAsk('Voulez-vous voir la liste des écritures détruites ?','')=mrYes then VisuPiecesGenere(ListePieces,EcrAna,0);
  end;
  VideListe(ListePieces);
  ListePieces.Free ;
  if ECran.Parent is THPanel then _DeblocageMonoPoste(False);
end ;

procedure TOF_CPMULSUPPRANA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULSUPPRANA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULSUPPRANA.Y_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(Y_EXERCICE.Value,Y_DATECOMPTABLE,Y_DATECOMPTABLE_);
end;

procedure TOF_CPMULSUPPRANA.Y_AXEChange(Sender: TObject);
begin
  if (Y_AXE.ItemIndex = 0) then Y_SECTION.DataType := 'TZSECTIONTOUS' else
  if (Y_AXE.ItemIndex = 1) then Y_SECTION.DataType := 'TZSECTION' else
                                Y_SECTION.DataType := 'TZSECTION'+IntToStr(Y_AXE.ItemIndex);
end;

procedure TOF_CPMULSUPPRANA.BDetruitesClick(Sender: TObject);
begin
  if ListePieces.Count>0 then VisuPiecesGenere(ListePieces,EcrAna,0) ;
end;

procedure TOF_CPMULSUPPRANA.FListeDblClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
  TrouveEtLanceSaisieODA(TFMul(Ecran).Q.TQ,taConsult);
{$ELSE}
  TrouveEtLanceSaisieODA(TFMul(Ecran).Q,taConsult);
{$ENDIF}
end;

procedure TOF_CPMULSUPPRANA.BOuvrirClick(Sender: TObject);
begin
  if DetruitSelection then TFMul(Ecran).BChercheClick(Nil);
end;

function TOF_CPMULSUPPRANA.DetruitSelection: boolean;
Var
  i,NbD : integer;
BEGIN
  Result := False;
  NbD := FListe.NbSelected;
  if NbD<=0 then begin
    PGIError('Vous n''avez sélectionné aucune écriture.', '');
    Exit;
  end;

  if PGIAsk('Confirmez-vous la suppression des écritures sélectionnées ?', '')<>mrYes then begin
    FListe.ClearSelected;
    Exit;
  end;

  Application.ProcessMessages;
  InitMove(NbD,'');

  for i:=0 to NbD-1 do begin
    FListe.GotoLeBookMark(i) ;
    {$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
    {$ENDIF}
    MoveCur(FALSE);
    if Transactions(DetruitLaPiece,5)<>oeOk then begin
      MessageAlerte('Certaines écritures en cours de traitement par un autre utilisateur n''ont pas été détruites.');
      FiniMove ;
      Exit ;
    end;
  end;
  FListe.ClearSelected;
  FiniMove;
  Result := True;
end;

procedure TOF_CPMULSUPPRANA.DetruitLaPiece;
var lTobODA : TOB ;
    i       : integer ;
    Q       : TQuery ;
begin
{$IFDEF EAGLCLIENT}
  Q := TFMul(Ecran).Q.TQ;
{$ELSE}
  Q := TFMul(Ecran).Q;
{$ENDIF}

  try

    lTobOda := Tob.Create('OLD_ODA', nil, -1) ;
    lTobOda.LoadDetailDBFromSQL('ANALYTIQ', 'SELECT * FROM ANALYTIQ '
                                 + ' WHERE Y_JOURNAL="' + Q.FindField('Y_JOURNAL').AsString + '"'
                                   + ' AND Y_DATECOMPTABLE="' + USDATETIME(Q.FindField('Y_DATECOMPTABLE').AsDateTime) + '"'
                                   + ' AND Y_QUALIFPIECE="N"'
                                   + ' AND Y_NUMEROPIECE=' + Q.FindField('Y_NUMEROPIECE').AsString);

    if lTobOda.Detail.count > 0 then
      begin

      for i:=lTobOda.Detail.count-1 downto 0 do
        begin
        // Maj Solde
        MajSoldeSectionTOB(lTobOda.Detail[i], False);
        // Maj journaux
        MajJournalAnaTob(lTobOda.Detail[i], False, False);
        // suppression mémoire sauf detail[0]
        if i>0 then
          lTobOda.Detail[i].Free ;
        end ;

      // Delete
      ExecuteSQL( 'DELETE FROM ANALYTIQ WHERE ' + WhereEcritureTOB( tsAnal, lTobOda.Detail[0], False) ) ;

      ListePieces.Add(lTobOda.Detail[0]);
      end ;

    except
      on E : Exception do
        PGIError(TraduireMemoire('Une erreur est survenue lors de la suppression de la pièce d''ODA :') +
                 #13#13 + E.Message, Ecran.Caption);

    end ;

end;

procedure TOF_CPMULSUPPRANA.InitCriteres;
begin
  // YM 25/08/2005    Initialisation des zones
  If Y_JOURNAL.Items.Count>0 then
  BEGIN
     If Y_JOURNAL.Vide then
         Y_JOURNAL.ItemIndex := 1
     Else
         Y_JOURNAL.ItemIndex := 0;
  END ;

  IF VH^.Suivant.Code<>'' then
      Y_EXERCICE.Value:=VH^.Suivant.Code
  else
      Y_EXERCICE.Value:=VH^.Encours.Code ;

  Y_AXE.ItemIndex := 0

end;

Initialization
  registerclasses ( [ TOF_CPMULSUPPRANA ] ) ;
end.
