{ Unité : Source TOF de la FICHE : TRMULICC
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.09.001.001  12/10/06  JP   Création de l'unité
 7.09.001.010  27/03/07  JP   FQ 10419 : on interdit la suppression des écritures intégrées en compta
 8.00.001.021  20/06/07  JP   FQ 10480 : gestion des concepts
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRMULICC_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1,
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  HTB97, UTob, Menus, ULibPieceCompta;


type
  {$IFDEF TRCONF}
  TOF_TRMULICC = class (TOFCONF)
  {$ELSE}
  TOF_TRMULICC = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  private
    TobCompta : TobPieceCompta;
    PopupMenu : TPopUpMenu;

    procedure ListeDbClick   (Sender : TObject);
    procedure NouveauClick   (Sender : TObject);
    procedure DeleteClick    (Sender : TObject);
    procedure RealiseClick   (Sender : TObject);
    procedure NoDossierChange(Sender : TObject);
  end ;


procedure TRLanceFiche_MulIcc(Arguments : string);


implementation


uses
  TRSAISIEFLUX_TOF, Constantes, Math, UProcGen, UProcSolde, Commun,
  ParamSoc, HQry, UProcCommission, UProcEcriture, ExtCtrls{TImage}, HMsgBox;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulIcc(Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  {Tester les droits !!!!!}
  AGLLanceFiche('TR', 'TRMULICC', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULICC.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000152;

  TFMul(Ecran).FListe.OnDblClick := ListeDbClick;
  TFMul(Ecran).BInsert.OnClick   := NouveauClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := DeleteClick;

  TToolbarButton97(GetControl('BREALISE')).OnClick := RealiseClick;
  TobCompta := TobPieceCompta.Create('****', nil, -1);

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.Items[0].OnClick := NouveauClick;
  PopupMenu.Items[1].OnClick := DeleteClick;
  PopupMenu.Items[3].OnClick := RealiseClick;
  AddMenuPop(PopupMenu, '', '');

  {20/06/07 : FQ 10480 : Gestion du concept la création / Suppresion des flux et VBO}
  CanCreateEcr (PopupMenu.Items[0]);
  CanCreateEcr (PopupMenu.Items[1]);
  CanValidateBO(PopupMenu.Items[3]);
  CanCreateEcr (TFMul(Ecran).BInsert);
  CanCreateEcr (GetControl('BDELETE'));
  CanValidateBO(GetControl('BREALISE'));

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, '');
  THValComboBox(GetControl('TE_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
  THValComboBox(GetControl('TE_NODOSSIER')).OnChange := NoDossierChange;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULICC.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobCompta) then FreeAndNil(TobCompta);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULICC.ListeDbClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  {$IFDEF EAGLCLIENT}
  if TFMul(Ecran).FListe.RowCount = 0 then Exit;
  {$ELSE}
  if TFMul(Ecran).Q.Eof and TFMul(Ecran).Q.Bof then Exit;
  {$ENDIF}
  s := VarToStr(GetField('TE_NODOSSIER'  )) + ';' + VarToStr(GetField('TE_NUMTRANSAC')) + ';' +
       VarToStr(GetField('TE_NUMEROPIECE')) + ';' + VarToStr(GetField('TE_NUMLIGNE'));

  {+ TRANSACICC : si l'on veut autoriser la modification des dates ... à voir !}
  AGLLanceFiche('TR', 'TRFICECRITURE', '', s, VarToStr(GetField('TE_NATURE')) + ';'{+ TRANSACICC + ';'});
end;

{Création d'une nouvelle écriture
{---------------------------------------------------------------------------------------}
procedure TOF_TRMULICC.NouveauClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', ';ICCMUL;');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{Suppression des écritures sélectionnées
{---------------------------------------------------------------------------------------}
procedure TOF_TRMULICC.DeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  s : string;
  l : TStringList;
  o : TObjDtValeur;
  MsgOk : Boolean;
begin
  if TFMul(Ecran).FListe.nbSelected = 0 then begin
    PGIBox(TraduireMemoire('Veuillez sélectionner au moins un flux.'), Ecran.Caption);
    Exit;
  end;

  {Liste mémorisant les comptes traités pour un recalcul des soldes}
  l := TStringList.Create;
  l.Duplicates := dupIgnore;
  l.Sorted := True;
  MsgOk := False;
  try
    if HShowMessage('0;Suppression d''écritures prévisionnelles;Êtes-vous sûr de vouloir supprimer les écritures sélectionnées ?;Q;YNC;N;C;', '', '') = mrYes then begin
      for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
        TFMul(Ecran).FListe.GotoLeBookmark(n);
        s := VarToStr(GetField('TE_GENERAL'));
        p := l.IndexOf(s);
        if p > -1 then
          {On récupère la date la plus ancienne pour le recalcul du compte}
          TObjDtValeur(l.Objects[p]).DateVal := Min(TObjDtValeur(l.Objects[p]).DateVal, VarToDateTime(GetField('TE_DATECOMPTABLE')))
        else begin
          o := TObjDtValeur.Create;
          o.DateVal := VarToDateTime(GetField('TE_DATECOMPTABLE'));
          l.AddObject(s, o);
        end;

        {27/03/07 : FQ 10419 : On interdit la suppression des écritures intégrées en compta}
        if VarToStr(GetField('TE_USERCOMPTABLE')) <> '' then begin
          if not msgOk then begin
            PGIBox(TraduireMemoire('Certains flux ont été intégrés en comptabilité et ne peuvent être supprimés'), Ecran.Caption);
            MsgOk := True;
          end;
          Continue;
        end;

        {20/07/04 : suppression des écritures avec la gestion des commissions}
        if not SupprimeEcriture(Ecran.Caption, VarToStr(GetField('TE_NUMTRANSAC')),
                                               VarToStr(GetField('TE_NUMEROPIECE')),
                                               VarToStr(GetField('TE_NUMLIGNE')),
                                               VarToStr(GetField('TE_NODOSSIER'))) then Exit;
      end;

      {Recalcul des soldes}
      for n := 0 to l.Count - 1 do
        RecalculSolde(l[n], DateToStr(TObjDtValeur(l.Objects[n]).DateVal), 0, True);
    end;

  finally
    LibereListe(l, True);
  end;

  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{JP 12/03/04 : réalisation des écritures et éventuelle intégration en comptabilité
{---------------------------------------------------------------------------------------}
procedure TOF_TRMULICC.RealiseClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  w : string;
  lTobEcr : Tob ;
begin
  if TFMul(Ecran).FListe.nbSelected = 0 then begin
    PGIBox(TraduireMemoire('Veuillez sélectionner au moins un flux.'), Ecran.Caption);
    Exit;
  end;

  if HShowMessage('1;' + Ecran.Caption + ';Voulez-vous réaliser les écritures sélectionnées ?;Q;YNC;N;N;', '', '') = mrYes then begin
    w := '';
    {1/ Réalisation}
    for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
      TFMul(Ecran).FListe.GotoLeBookmark(n);
      if GetField('TE_NATURE') = na_Prevision then begin
         w :=  w + '"' + GetField('TE_NUMTRANSAC') + '",';
      end;
    end;

    if w = '' then begin
      HShowMessage('0;' + Ecran.Caption + ';Aucune écriture à réaliser.;I;O;O;O;', '', '');
      Exit;
    end;

    {Suppression de la dernière virgule}
    System.Delete(w, Length(w), 1);
    {Réalisation à proprement parler}
    ExecuteSQL('UPDATE TRECRITURE SET TE_NATURE = "R" WHERE TE_NUMTRANSAC IN (' + w + ')');

    {2/ Intégration en compta}
    if not GetParamSocSecur('SO_TRINTEGAUTO', False) and not AutoriseFonction(dac_Integration) then Exit;
    if HShowMessage('2;' + Ecran.Caption + ';Voulez-vous intégrer en comptabilité les écritures réalisées ?;Q;YNC;N;N;', '', '') = mrNo then Exit;

    {27/05/05 : Nouvelle gestion des erreurs}
    InitGestionErreur;
    CategorieCurrent := CatErr_CPT;

    {Chargement en tob des écritures à intégrer}
    lTobEcr := TOb.Create('$TOBECR', nil, -1) ;
    lTobEcr.LoadDetailFromSQL( 'SELECT * FROM TRECRITURE WHERE TE_NUMTRANSAC IN (' + w + ') ORDER BY TE_NODOSSIER, TE_NUMTRANSAC, TE_COMMISSION, TE_NUMEROPIECE, TE_NUMLIGNE' ) ;

    try
      {Générations des écritures de comptabilité à partir du record}
      TRGenererPieceCompta( TobCompta, lTobEcr );

      {11/01/06 : FQ 10323 : Vérifie si les écritures sont bonnes}
      AfficheMessageErreur(Ecran.Caption);

      {Maintenant que l'on a les écritures au format de la compta dans la tobCompta,
       on lance le processus d'intégration en comptabilité proprement dit}
      if TobCompta.Detail.Count > 0 then
        {21/08/06 : Nouvelle fonction de traitement}
        TRIntegrationPieces(TobCompta, False );

    finally
      if Assigned(lTobEcr) then
        FreeAndNil(lTobEcr);
    end;

  end;

  TFMul(Ecran).FListe.ClearSelected;
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULICC.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL', '');
end;


initialization
  RegisterClasses([TOF_TRMULICC]);

end.
