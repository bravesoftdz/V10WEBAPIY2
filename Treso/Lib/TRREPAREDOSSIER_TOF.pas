unit TRREPAREDOSSIER_TOF;
{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
7.06.001.001   07/11/06   JP   Création de l'unité : Elle a pour but de réparer certains dossiers,
                               en particulier s'il y a eu un plantage dans lors de la mise en place
                               du Multi sociétés dans la Trésorerie

--------------------------------------------------------------------------------------}
interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_MAIN,
  {$ENDIF}
  StdCtrls, Controls, Classes, HCtrls, HEnt1, UTOF;


type
  TOF_TRREPAREDOSSIER = class (TOF)
    procedure OnArgument(S : string); override;
    procedure OnUpdate              ; override;
  private
    procedure InitControls;
    procedure FiltreControls;

    procedure MajNoDossier;
    procedure MajSoldes;
  public
    procedure HelpNodossierClick(Sender : TObject);
    procedure HelpReinitClick   (Sender : TObject);

    procedure CkNodossierClick  (Sender : TObject);
    procedure CkReinitClick    (Sender : TObject);
  end;

procedure TRLanceFiche_PrepareDossier(Arguments : string);


implementation

uses
  {$IFDEF EAGLCLIENT}
  uTob,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  SysUtils, HMsgBox, HTB97, URecupDos, LicUtil, Commun, Constantes, UProcGen;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_PrepareDossier(Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  if V_PGI.Superviseur and (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    AGLLanceFiche('TR', 'TRREPAREDOSSIER', '', '', Arguments)
  else
    HShowMessage('0;Réparation des dossiers;Pour accéder à ce traitement, il faut être superviseur et '#13 +
                 'entrer dans la Trésorerie avec le mot de passe du jour;W;O;O;O;', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  InitControls;
  FiltreControls;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if TCheckBox(GetControl('CKMAJNODOSSIER')).Checked then MajNoDossier;
  if TCheckBox(GetControl('CKMAJREINIT'   )).Checked then MajSoldes;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.InitControls;
{---------------------------------------------------------------------------------------}
begin
  TToolbarButton97(GetControl('HLPNODOSSIER')).OnClick := HelpNodossierClick;
  TToolbarButton97(GetControl('HLPREINIT'   )).OnClick := HelpReinitClick;
  TCheckBox(GetControl('CKMAJNODOSSIER'     )).OnClick := CkNodossierClick;
  TCheckBox(GetControl('CKMAJREINIT'        )).OnClick := CkReinitClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.FiltreControls;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := 'DOS_NOMBASE ' + FiltreBaseTreso(False);
  THMultiValComboBox(GetControl('CBNODOSSIER')).Plus := s;
  THMultiValComboBox(GetControl('CBREINIT')).Plus := FiltreBanqueCp(THMultiValComboBox(GetControl('CBNODOSSIER')).DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.CkNodossierClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlEnabled('CBNODOSSIER', TCheckBox(Sender).Checked);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.CkReinitClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlEnabled('CBREINIT', TCheckBox(Sender).Checked);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.HelpNodossierClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  HShowMessage('0;Mise à jour du Numéro de dossier;Ce traitement à pour but de mettre à jour les champs'#13 +
               '"XX_NODOSSIER" sur les tables de mouvements (TRECRITURE, EQUILIBRAGE,'#13 +
               'COURTSTERMES, TROPCVM, TRVENTEOPCVM) pour les dossiers sélectionnés.;I;O;O;O;', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.HelpReinitClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  HShowMessage('0;Mise à jour des soldes de Réinitialisation;Ce traitement à pour but de recalculer les soldes des'#13 +
               'comptes sélectionnés après une réinitialisation des soldes au'#13 +
               '1er janvier au matin.;I;O;O;O;', '', '');
end;


{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.MajNoDossier;
{---------------------------------------------------------------------------------------}
begin
  if HShowMessage('0;' + Ecran.Caption + ';;Q;YN;N;N', '', '') = mrNo then Exit;
  if HShowMessage('0;' + Ecran.Caption + ';;Q;YN;Y;Y', '', '') = mrYes then Exit;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRREPAREDOSSIER.MajSoldes;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TQuery;
begin
  if HShowMessage('0;' + Ecran.Caption + ';Confirmez-vous la réinitialisation des soldes ?;Q;YN;N;N', '', '') = mrNo then Exit;
  if HShowMessage('0;' + Ecran.Caption + ';Souhaitez-vous abandonner la réinitialisation des soldes ?;Q;YN;Y;Y', '', '') = mrYes then Exit;
  {S'il n'y a pas de filtre sur les comptes, on les traite}
  if THMultiValComboBox(GetControl('CBREINIT')).Tous or
     (GetControlText('CBREINIT') = '') then
    MajReInit
  else begin
    SQL := GetControlText('CBREINIT');
    SQL := GetClauseIn(SQL);
    SQL := 'SELECT BQ_CODE FROM BANQUECP WHERE BQ_CODE IN (' + SQL + ')';
    Q := OpenSQL(SQL, True);
    try
      while not Q.EOF do begin
        MajReInit(Q.FindField('BQ_CODE').AsString);
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;
  end;
end;

initialization
  RegisterClasses([TOF_TRREPAREDOSSIER]);

end.
