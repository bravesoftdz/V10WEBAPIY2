{ Unité : Source TOF du Mul : CPMULAFFECTBAP ()
--------------------------------------------------------------------------------------
    Version   | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  02/02/06  JP   Création de l'unité
 8.01.001.003  26/01/07  JP   FQ 19451 : Mise à jour des dates en fonction du choix de l'exercice
 8.10.002.001  15/10/07  JP   FQ 16149 : gestion des restrictions utilisateur
--------------------------------------------------------------------------------------}
unit CPMULAFFECTBAP_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eMul,
  {$ELSE}
  FE_Main, db, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} Mul, HDB,
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  CPANCETREMUL, {JP 15/10/07 : FQ 16149 : création d'un ancêtre}
  uLibBonAPayer, SysUtils, UTob, ComCtrls, uTOF, HCtrls;

type
  TOF_CPMULAFFECTBAP = class (TOFANCETREMUL)
    procedure OnLoad                ; override;
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnDisplay             ; override;
  private
    EnregBap : TBonAPayer;
    ListeErr : TList;

    procedure GenereBAP;
  public
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF EAGLCLIENT}

    procedure BOuvrirClick(Sender : TObject);
    procedure BAffectClick(Sender : TObject);
    procedure BManuelClick(Sender : TObject);
    procedure SlctAllClick(Sender : TObject);
    procedure ExoOnChange (Sender : TObject); {26/01/07 : FQ 19451} 
  end;

procedure CpLanceFiche_AffectationCircuit;

implementation

uses
  HEnt1, HMsgBox, Ent1, HTB97, CPGENEREBAPCREATE_TOF, CPGENEREBAP_TOF, AGLInit,
  UProcGen, ULibPieceCompta, ULibEcriture, CPSAISIEPIECE_TOF, ULibExercice;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_AffectationCircuit;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMULAFFECTBAP', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.OnArgument(S: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509050;
  FListe := TFMul(Ecran).FListe;
  TToolbarButton97(GetControl('BOUVRIR'   )).OnClick := BOuvrirClick;
  TToolbarButton97(GetControl('BAUTO'     )).OnClick := BAffectClick;
  TToolbarButton97(GetControl('BMANUEL'   )).OnClick := BManuelClick;
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;
  FListe.OnDblClick := BOuvrirClick;
  SetPlusCombo(THValComboBox(GetControl('E_JOURNAL')), 'J');
  {JP 26/01/07 : FQ 19451 : Mise à jour des dates en fonction du choix de l'exercice}
  (GetControl('E_EXERCICE') as THValComboBox).OnChange := ExoOnChange;
  ExoOnChange(GetControl('E_EXERCICE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(EnregBap) then FreeAndNil(EnregBap);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.OnDisplay;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.SlctAllClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.BOuvrirClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  lStJournal  : string ;
  lDtDateC    : TDateTime ;
  lInNumPiece : integer ;
begin
  // SBO 02/10/2007 FQ 19174 : Remplacement de la fonction ModifiePieceCompta par SaisiePieceCompta plus standard
  lStJournal  := GetField('E_JOURNAL') ;
  lDtDateC    := GetField('E_DATECOMPTABLE') ;
  lInNumPiece := GetField('E_NUMEROPIECE') ;
  SaisiePieceCompta( 'N', taConsult, lStJournal, lDtDateC, lInNumPiece ) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.BManuelClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  S : string;
begin
  S := 'N;' + VarToStr(GetField('E_JOURNAL')) + ',' + VarToStr(GetField('E_EXERCICE')) + ',' +
       DateToStr(VarToDateTime(GetField('E_DATECOMPTABLE'))) + ',' + VarToStr(GetField('E_NUMEROPIECE')) + ',;';
  CpLanceFiche_GenereBap(S);
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.BAffectClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}
  n : Integer;
begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end
  else if HShowMessage('0;' + Ecran.Caption + ';Êtes vous sûr de générer des bons à payer sur les écritures sélectionnées ?;Q;YN;N;N;', '', '') <> mrYes then
    Exit;

  ListeErr := TList.Create;
  try
    {$IFNDEF EAGLCLIENT}
    TFMul(Ecran).Q.First;
    if F.AllSelected then
      while not TFMul(Ecran).Q.EOF do begin
        GenereBAP;
        TFMul(Ecran).Q.Next;
      end
    else
    {$ENDIF}

    for n := 0 to F.nbSelected - 1 do begin
      F.GotoLeBookmark(n);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
      {$ENDIF}
      GenereBAP;
    end;

    EnregBap.TobMasse.InsertDb(nil);
  finally
    {...Affichage du résultat}
    TheData := ListeErr;
    TheTob  := EnregBap.TobMasse;
    CpLanceFiche_GenereBapColl;
    EnregBap.TobMasse.ClearDetail;
    if Assigned(ListeErr) then DisposeListe(ListeErr, True);
  end;

  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.GenereBAP;
{---------------------------------------------------------------------------------------}
var
  RecC : TClefPiece;
  PRec : PClefPiece;
begin
  RecC.E_NUMEROPIECE   := TFMul(Ecran).Q.FindField('E_NUMEROPIECE'  ).AsInteger;
  RecC.E_DATECOMPTABLE := TFMul(Ecran).Q.FindField('E_DATECOMPTABLE').AsDateTime;
  RecC.E_EXERCICE      := TFMul(Ecran).Q.FindField('E_EXERCICE'     ).AsString;
  RecC.E_JOURNAL       := TFMul(Ecran).Q.FindField('E_JOURNAL'      ).AsString;
  RecC.E_QUALIFPIECE   := 'N';
  if not Assigned(EnregBap) then
    EnregBap := TBonAPayer.Create(RecC, True, True)
  else
    EnregBap.CreerBap(RecC);

  if EnregBap.BapImpossible then begin
    System.New(PRec);
    PRec^.E_JOURNAL := RecC.E_JOURNAL;
    PRec^.E_DATECOMPTABLE := RecC.E_DATECOMPTABLE;
    PRec^.E_NUMEROPIECE := RecC.E_NUMEROPIECE;
    PRec^.E_EXERCICE := RecC.E_EXERCICE;
    PRec^.E_QUALIFPIECE := GetMessageErreur(EnregBap.CodeErreur);
    ListeErr.Add(PRec);
  end;
end;

{FQ 19451 : Mise à jour des dates en fonction du choix de l'exercice
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULAFFECTBAP.ExoOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if GetControlText('E_EXERCICE') = '' then begin
    SetControlText('E_DATECOMPTABLE', StDate1900);
    SetControlText('E_DATECOMPTABLE_', StDate2099);
  end
  else begin
    SetControlText('E_DATECOMPTABLE', DateToStr(CtxExercice.QuelExoDate(GetControlText('E_EXERCICE')).Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(CtxExercice.QuelExoDate(GetControlText('E_EXERCICE')).Fin));
  end;
end;

initialization
  RegisterClasses([TOF_CPMULAFFECTBAP]);

end.

