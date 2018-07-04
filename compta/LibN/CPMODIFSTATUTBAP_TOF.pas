{ Unité : Source TOF de la fiche CPMODIFSTATUTBAP
--------------------------------------------------------------------------------------
    Version   |  Date  |Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  27/04/06  JP   Création de l'unité
 8.00.002.002  02/08/07  JP   Remplacement du THRichEditOle par un TMemo
--------------------------------------------------------------------------------------}
unit CPMODIFSTATUTBAP_TOF;

interface

uses
  Controls, Classes, Vierge,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  uLibBonAPayer, SysUtils, HCtrls, UTob, HEnt1, UTOF, StdCtrls;

type
  TOF_CPMODIFSTATUTBAP = class(TOF)

    procedure OnArgument(S : string); override;
    procedure OnUpdate             ; override;
  private
    FBlocNote : TMemo;
    ClePiece  : TClefPiece;

    procedure AfficheBap(Clef : string);

    procedure BOuvrirPieceClick(Sender : TObject);
    procedure BOuvrirBAPClick  (Sender : TObject);
  end;

procedure CpLanceFiche_ModifStatutBap(Clef : string);

implementation

uses
  ULibPieceCompta, ULibEcriture, CPSAISIEPIECE_TOF, HTB97, HMsgBox,
  CPBONSAPAYER_TOF;


{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_ModifStatutBap(Clef : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMODIFSTATUTBAP', '', '', Clef + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFSTATUTBAP.AfficheBap(Clef : string);
{---------------------------------------------------------------------------------------}
var
  Q    : TQuery;
begin
  ClePiece := DecodeClefTP(Clef, ',');
  Q := OpenSQL('SELECT BAP_BLOCNOTE, BAP_STATUTBAP FROM CPBONSAPAYER WHERE ' +
                    'BAP_JOURNAL = "' + ClePiece.E_JOURNAL + '" AND ' +
                    'BAP_EXERCICE = "' + ClePiece.E_EXERCICE + '" AND ' +
                    'BAP_DATECOMPTABLE = "' + USDateTime(ClePiece.E_DATECOMPTABLE) + '" AND ' +
                    'BAP_NUMEROPIECE = ' + IntToStr(ClePiece.E_NUMEROPIECE) + 'AND ' +
                    'BAP_NUMEROORDRE = ' + IntToStr(ClePiece.E_NUMLIGNE), True);
  try
    if not Q.EOF then begin
      SetControlText('CBJOURNAL', ClePiece.E_JOURNAL);
      SetControlText('CBEXERCICE', ClePiece.E_EXERCICE);
      SetControlText('EDDATE', DateToStr(ClePiece.E_DATECOMPTABLE));
      SetControlText('EDNUMEROPIECE', IntToStr(ClePiece.E_NUMEROPIECE));
      SetControlText('CBSTATUT', Q.FindField('BAP_STATUTBAP').AsString);
      FBlocNote.Lines.Text := Q.FindField('BAP_BLOCNOTE').AsString;
    end
    else begin
      HShowMessage('0;' + Ecran.Caption + ';L''étape est introuvable;E;O;O;O;', '', '');
      SetControlEnabled('BPIECE'       , False);
      SetControlEnabled('BSTATUT'      , False);
      SetControlEnabled('RECOMMENTAIRE', False);
      SetControlEnabled('BVALIDER'     , False);
    end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFSTATUTBAP.BOuvrirPieceClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  lInfoEcr  : TInfoEcriture;
  lPieceCpt : TPieceCompta;
begin
  {Chargement pièce}
  lInfoEcr  := TInfoEcriture.Create(V_PGI.SchemaName);
  lPieceCpt := TPieceCompta.CreerPiece(lInfoEcr);
  try
    {Champs obligatoires...}
    lPieceCpt.PutEntete('E_JOURNAL',       GetControlText('CBJOURNAL'));
    lPieceCpt.PutEntete('E_EXERCICE',      GetControlText('CBEXERCICE'));
    lPieceCpt.PutEntete('E_DATECOMPTABLE', StrToDate(GetControlText('EDDATE')));
    lPieceCpt.PutEntete('E_NUMEROPIECE',   ValeurI(GetControlText('EDNUMEROPIECE')));
    lPieceCpt.PutEntete('E_QUALIFPIECE',   'N');

    {Chargement}
    lPieceCpt.LoadFromSQL ;

    {Ouverture Saisie}
    ModifiePieceCompta(lPieceCpt, taConsult);
  finally
    FreeAndNil(lInfoEcr);
    FreeAndNil(lPieceCpt);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFSTATUTBAP.BOuvrirBAPClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_FicheBap(GetControlText('CBJOURNAL') + ',' + GetControlText('CBEXERCICE') + ',' +
                        GetControlText('EDDATE') + ',' + GetControlText('EDNUMEROPIECE') + ',;');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFSTATUTBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  inherited;
  FBlocNote := TMemo(GetControl('RECOMMENTAIRE'));

  {Récupération de la Clef}
  ch := ReadTokenSt(S);
  AfficheBap(ch);

  TToolBarButton97(GetControl('BPIECE' )).OnClick := BOuvrirPieceClick;
  TToolBarButton97(GetControl('BSTATUT')).OnClick := BOuvrirBAPClick;
  Ecran.HelpContext := 7509085;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFSTATUTBAP.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  TobBap : TOB;
begin
  inherited;
  if HShowMessage('0;' + Ecran.Caption + ';Êtes-vous sûr de vouloir modifier le statut du bon à'#13 +
                  'payer et enregistrer les modifications des commentaires ?;Q;YN;Y;N;', '', '') = mrYes then begin
    TobBap  := TOB.Create('CPBONSAPAYER', nil, -1);
    try
      TobBap.SetString  ('BAP_JOURNAL'      , ClePiece.E_JOURNAL);
      TobBap.SetString  ('BAP_EXERCICE'     , ClePiece.E_EXERCICE);
      TobBap.SetDateTime('BAP_DATECOMPTABLE', ClePiece.E_DATECOMPTABLE);
      TobBap.SetInteger ('BAP_NUMEROPIECE'  , ClePiece.E_NUMEROPIECE);
      TobBap.SetInteger ('BAP_NUMEROORDRE'  , ClePiece.E_NUMLIGNE);
      TobBap.LoadDB;
      if TobBap.GetString('BAP_STATUTBAP') <> '' then begin
        TobBap.SetDateTime('BAP_DATEMODIF', Now);
        TobBap.SetString('BAP_MODIFICATEUR', V_PGI.User);
        TobBap.SetString('BAP_STATUTBAP', sbap_Encours);
        TobBap.PutValue('BAP_BLOCNOTE', FBlocNote.Lines.Text);
        TobBap.UpdateDB;
      end
      else
        HShowMessage('0;' + Ecran.Caption + ';Impossible de mettre à jour le BAP.;E;O;O;O;', '', '');
    finally
      FreeAndNil(TobBap);
    end;
  end;
end;

initialization
  RegisterClasses([TOF_CPMODIFSTATUTBAP]);

end.

