{ Unité : Source TOF de la fiche CPGENEREBAP
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  08/02/06  JP   Création de l'unité
 8.00.001.001  07/12/06  JP   FQ 19163 : Afficher les libellé des user plutôt que les codes
 8.00.001.020  14/06/07  JP   FQ 20551 : rafraîchissement de la tablette sur les types de visa
--------------------------------------------------------------------------------------}
unit CPGENEREBAP_TOF;

interface

uses
  Controls, Classes, Vierge, 
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  Grids, SysUtils, HCtrls, UTob, HEnt1, uLibBonAPayer, UTOF, Forms;

type
  TActionBap = (ab_Consultion, ab_Creation, ab_Lien);

  TOF_CPGENEREBAP = class(TOF)
    FListe : THGrid;

    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
  private
    BapAction  : TActionBap;
    EnregBap   : TBonAPayer;
    NbBap      : Integer;
    PeutFermer : Boolean;
    FCanClose  : TCloseQueryEvent;

    procedure AfficheBap   (Clef : string);
    procedure ChargerGrille(Code : string);
    procedure ChargeBap    ; 
  public
    RecC : TClefPiece;

    procedure CircuitChange  (Sender : TObject);
    procedure TypeVisaChange (Sender : TObject);
    procedure BParamTypeClick(Sender : TObject);{JP 14/06/07 : FQ 20551}
    procedure ECanClose      (Sender : TObject; var CanClose : Boolean);
    procedure BFermeMouseDown(Sender : TObject; Button : TMouseButton; Shift : TShiftState; X, Y : Integer);
  end;

procedure CpLanceFiche_GenereBap(Argument : string);

implementation

uses
  CPTYPEVISA_TOM, {JP 14/06/07 : FQ 20551}
  HMsgBox, UProcGen, HTB97;


{Arg 1 : N (Création d'un Bap) / C (Consultation d'un Bap) /
         L (Consultation du BAP d'une écriture
 Arg 2 : Clef : JOURNAL,EXERCICE,DATECOMPTABLE,NUMEROPIECE,
{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_GenereBap(Argument : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPGENEREBAP', '', '', Argument);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  inherited;
  Ecran.HelpContext := 7509040;
  NbBap := 0;
  FListe := THGrid(GetControl('FLISTE'));

  ch := ReadTokenSt(S);
  case StrToChr(ch) of
    'C' : BapAction := ab_Consultion;
    'L' : BapAction := ab_Lien;
    'N' : BapAction := ab_Creation;
    else
      BapAction := ab_Consultion;
  end;

  {Récupération de la Clef}
  ch := ReadTokenSt(S);
  AfficheBap(ch);
  THValComboBox(GetControl('CBCIRCUIT' )).OnChange := CircuitChange;
  THValComboBox(GetControl('CBCODEVISA')).OnChange := TypeVisaChange;

  {Pour Bloquer la fermeture de la fiche si la saisie n'est pas correcte}
  TToolBarButton97(GetControl('BFERME')).OnMouseDown := BFermeMouseDown;
  {JP 14/06/07 : FQ 20551 : On ne passe pas par DataTypeParametrable pour avoir la main}
  TToolBarButton97(GetControl('BPARAMTYPE')).OnClick := BParamTypeClick;

  FCanClose := Ecran.OnCloseQuery;
  Ecran.OnCloseQuery := ECanClose;
  PeutFermer := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(EnregBap) then FreeAndNil(EnregBap);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited;
  if BapAction = ab_Creation then begin
    if GetControlText('CBCODEVISA') = '' then begin
      LastError := 3;
      LastErrorMsg := TraduireMemoire('Veuillez sélectionner un type de visa');
      PeutFermer := False;
    end
    else if GetControlText('CBCIRCUIT') = '' then begin
      LastError := 2;
      LastErrorMsg := TraduireMemoire('Veuillez sélectionner un circuit');
      PeutFermer := False;
    end;

    if LastErrorMsg = '' then begin
      if EnregBap.BapImpossible then
        EnregBap.ForceBap(RecC, GetControlText('CBCODEVISA'), GetControlText('CBCIRCUIT'));

      T := EnregBap.TobBap.FindFirst(['BAP_CODEVISA'], [GetControlText('CBCODEVISA')], True);
      if not Assigned(T) then begin
        LastError := 4;
        LastErrorMsg := TraduireMemoire('Impossible de charger le bons à payer');
        PeutFermer := False;
      end
      else begin
        {Le circuit a été modifié on met à jour le BAP avant de mettre en base}
        if T.GetString('BAP_CIRCUITBAP') <> GetControlText('CBCIRCUIT') then begin
          T.SetString('BAP_CIRCUITBAP', GetControlText('CBCIRCUIT'));
          EnregBap.GetCircuit(1, T);
        end;

        T.InsertDb(nil);
        PeutFermer := True;
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.ChargerGrille(Code : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  s : string; {FQ 19163} 
begin
  T := TOB.Create('AACPCIRCUIT', nil, -1);
  try
    FListe.VidePile(False);
    {08/12/06 : FQ 19163 : jointure sur Utilisat}
    S := 'SELECT CCI_NUMEROORDRE, CCI_CIRCUITBAP, U1.US_LIBELLE US1, U2.US_LIBELLE US2, CCI_NBJOUR FROM CPCIRCUIT ';
    S := S + 'LEFT JOIN UTILISAT U1 ON CCI_VISEUR1 = U1.US_UTILISATEUR ';
    S := S + 'LEFT JOIN UTILISAT U2 ON CCI_VISEUR2 = U2.US_UTILISATEUR ';
    S := S + 'WHERE CCI_CIRCUITBAP = "' + Code + '" ORDER BY CCI_NUMEROORDRE';
    T.LoadDetailFromSQL(S);
    T.PutGridDetail(FListe, False, False, 'CCI_NUMEROORDRE;CCI_CIRCUITBAP;US1;US2;CCI_NBJOUR');
  finally
    FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.CircuitChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  ChargerGrille(GetControlText('CBCIRCUIT'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.TypeVisaChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  P : string;
  S : string;
  n : Integer;
begin
  T := TOB.Create('CPTYPEVISA', nil, -1);
  try
    T.SetString('CTI_CODEVISA', GetControlText('CBCODEVISA'));
    T.LoadDB;
    T.PutEcran(Ecran, TWinControl(GetControl('PNCRITERE')));
  finally
    FreeAndNil(T);
  end;

  if GetControlText('CTI_NBVISA') > '-1' then
    P := 'AND CCI_CIRCUITBAP IN (SELECT CCI_CIRCUITBAP FROM CPVUECIRCUIT ' +
         'WHERE CCI_NBVISAS = ' + GetControlText('CTI_NBVISA') + ') ';

  {S'il y a plusieurs BAP, On limite les circuits à ceux figurants dans l'objet}
  if NbBap > 0 then begin
    S := 'AND CCI_CIRCUITBAP IN (';
    for n := 0 to EnregBap.TobBap.Detail.Count - 1 do
      S := S + '"' + EnregBap.TobBap.Detail[n].GetString('BAP_CIRCUITBAP') + '",';
    {On écrase la dernière virgule par une parenthèse}
    S[Length(S)] := ')';
    P := P + S;
  end;

  THValComboBox(GetControl('CBCIRCUIT')).Plus := P;
  THValComboBox(GetControl('CBCIRCUIT')).Refresh;
  T := EnregBap.TobBap.FindFirst(['BAP_CODEVISA'], [GetControlText('CBCODEVISA')], True);
  if Assigned(T) then SetControlText('CBCIRCUIT', T.GetString('BAP_CIRCUITBAP'))
                 else THValComboBox(GetControl('CBCIRCUIT')).ItemIndex := 0;
  CircuitChange(GetControl('CBCIRCUIT'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.AfficheBap(Clef : string);
{---------------------------------------------------------------------------------------}
begin
  RecC := DecodeClefTP(Clef, ',');
  if RecC.E_QUALIFPIECE = '' then RecC.E_QUALIFPIECE := 'N';
  {14/06/07 : FQ 20551 : Le code a été déplacé dans ChargeBap afin de pouvoir être rappelé
              à la suite d'une modification d'un type de visa}
  ChargeBap;

  {Affichage du  premier type de visa}
  if not EnregBap.BapImpossible then begin
    SetControlText('CBCODEVISA', EnregBap.TobBap.Detail[0].GetString('BAP_CODEVISA'));
    TypeVisaChange(GetControl('CBCODEVISA'));
    SetControlText('CBCIRCUIT' , EnregBap.TobBap.Detail[0].GetString('BAP_CIRCUITBAP'));
    ChargerGrille(GetControlText('CBCIRCUIT'));
  end;

  if BapAction <> ab_Creation then begin
    SetControlEnabled('PNHAUT', False);
    Ecran.Caption := TraduireMemoire('Consultation d''un bon à payer');
    UpdateCaption(Ecran);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.ChargeBap;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  S : string;
begin
  if BapAction = ab_Creation then EnregBap := TBonAPayer.Create(RecC, True)
                             else EnregBap := TBonAPayer.Create(RecC, False);

  if EnregBap.BapImpossible then begin
    HShowMessage('0;' + Ecran.Caption + ';' + GetMessageErreur(EnregBap.CodeErreur) + ';W;O;O;O;', '', '');
    NbBap := 0;
  end
  else begin
    {Filtre sur les types de visas en fonction de la pièce que l'on traite}
    NbBap := EnregBap.TobBap.Detail.Count;
    S := 'CTI_CODEVISA IN (';
    for n := 0 to EnregBap.TobBap.Detail.Count - 1 do
      S := S + '"' + EnregBap.TobBap.Detail[n].GetString('BAP_CODEVISA') + '",';
    {On écrase la dernière virgule par une parenthèse}
    S[Length(S)] := ')';
    THValComboBox(GetControl('CBCODEVISA')).Plus := S;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.ECanClose(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := PeutFermer;
  if CanClose then FCanClose(Sender, Canclose);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.BFermeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
begin
  PeutFermer := True;
end;

{JP 14/06/07 : FQ 20551 : On ne passe pas par DataTypeParametrable pour avoir la main
{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAP.BParamTypeClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_TypeVisaFiche('', GetControlText('CBCODEVISA'), '');
  if Assigned(EnregBap) then FreeAndNil(EnregBap);
  ChargeBap;
end;

initialization
  RegisterClasses([TOF_CPGENEREBAP]);

end.
