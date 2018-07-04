{ Unité : Source TOF de la Fiche : TRINTEGRECOMPTA
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91          21/07/03  JP   Création de l'unité
 1.0.0.001.005 24/02/04  JP   Correction d'un sur la clause Where sur les journaux
 6.0x.xxx.xxx  05/08/04  JP   Gestion des commissions dans l'intégration
 6.50.001.001  27/05/05  JP   FQ 10223 : Nouvelle gestion des erreurs
 6.50.001.003  08/06/05  JP   Dans le cadre de la FQ 10223, Refonte de l'intégration en compta
 7.00.001.001  12/01/06  JP   FQ 10323 : Correction de la gestion de la TVA
 7.05.001.001  21/08/06  SBO  Intégration en compta multi-sociétés
 7.05.001.001  20/10/06  JP   Gestion des filtres sur les comptes et les dossiers
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRINTEGRECOMPTA_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eMul, Grids, MaineAGL,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Mul, DBGrids, FE_Main,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, UTOB, HTB97, ExtCtrls, Commun,
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  Constantes, UObjGen, ULibPieceCompta;

type
  {$IFDEF TRCONF}
  TOF_TRINTEGRECOMPTA = class (TOFCONF)
  {$ELSE}
  TOF_TRINTEGRECOMPTA = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument (S : string ) ; override ;
    procedure OnClose                  ; override ;
  private
    Fiche        : TFMul;
    {$IFDEF EAGLCLIENT}
    cbDevise : THValComboBox;
    {$ELSE}
    cbDevise : THDBValComboBox;
    {$ENDIF}

    procedure Exporter       (Sender : TObject);
    procedure AffFiche       (Sender : TObject);
    procedure DeviseChg      (Sender : TObject);
    procedure NoDossierChange(Sender : TObject);
  public
    TobGen : TobPieceCompta;
  end ;

  procedure TRLanceFiche_Integration(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  UProcEcriture, HStatus;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Integration(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRINTEGRECOMPTA.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000132;
  Fiche  := TFMul(Ecran);

  Fiche.FListe.OnDblClick := AffFiche;

  TToolbarButton97(GetControl('BOUVRIR')).OnClick := Exporter;
  {En attendant de le gérer}
  SetControlVisible('BSELECTALL', False);
  {$IFDEF EAGLCLIENT}
  cbDevise := THValComboBox(GetControl('TE_DEVISE'));
  {$ELSE}
  cbDevise := THDBValComboBox(GetControl('TE_DEVISE'));
  {$ENDIF}
  cbDevise.OnChange := DeviseChg;

  {Tob pour le processus d'intégration en comptabilité}
  TobGen := TobPieceCompta.Create('§ECRIT', nil, -1);

  {20/10/06 : gestion du multi sociétés}
  SetControlVisible('TE_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTE_NODOSSIER', IsTresoMultiSoc);

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL')).DataType, '', '');
  THValComboBox(GetControl('TE_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
  THValComboBox(GetControl('TE_NODOSSIER')).OnChange := NoDossierChange;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRINTEGRECOMPTA.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobGen) then FreeAndNil(TobGen);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRINTEGRECOMPTA.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL')).DataType, '', GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL', '');
end;

{Lancement de la fiche de l'écriture de trésorerie sélectionnée
{---------------------------------------------------------------------------------------}
procedure TOF_TRINTEGRECOMPTA.AffFiche(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  {$IFDEF EAGLCLIENT}
  if Fiche.FListe.RowCount = 0 then Exit;
  {$ELSE}
  if Fiche.Q.Eof and TFMul(Ecran).Q.Bof then Exit;
  {$ENDIF}
  s := GetField('TE_NODOSSIER') + ';' + GetField('TE_NUMTRANSAC') + ';' + VarToStr(GetField('TE_NUMEROPIECE')) + ';' +
       VarToStr(GetField('TE_NUMLIGNE'));
  AGLLanceFiche('TR', 'TRFICECRITURE', '', s, GetField('TE_NATURE') + ';');
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRINTEGRECOMPTA.DeviseChg(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), cbDevise.Value);
end ;

{Lancement du processus d'intégration en comptabilité
{---------------------------------------------------------------------------------------}
procedure TOF_TRINTEGRECOMPTA.Exporter(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  i, n : Integer;
  lTobEcr : Tob ;
begin
  if Fiche.FListe.VisibleRowCount = 0 then begin
    MessageAlerte(TraduireMemoire('Aucun élément à traité'));
    Exit;
  end;

  n := Fiche.FListe.nbSelected;

  if (n = 0) then begin
    MessageAlerte(TraduireMemoire('Aucun élément sélectionné'));
    exit;
  end;

  {Demande de confirmation de l'intégration}
  if HShowMessage('1;' + Ecran.Caption + ';Confirmez vous l''intégration en comptabilité des écritures sélectionnées?;Q;YN;N;N', '', '') <> mrYes then Exit;

  {FQ 10223 : Nouvelle gestion des erreurs}
  InitGestionErreur(CatErr_CPT);
  TobGen.ClearDetailPC;

  {Boucle sur les lignes sélectionnées}
  InitMove(n, '');
  try
    for i := 0 to n - 1 do begin
      Fiche.FListe.GotoLeBookmark(i);
      MoveCur(False);
      {... on prépare la Tob qui va conténir les écritures au format comptable}

      {Chargement en tob des écritures à intégrer}
      lTobEcr := TOb.Create('$TOBECR', nil, -1) ;
      lTobEcr.LoadDetailFromSQL( 'SELECT * FROM TRECRITURE '
                                        + 'WHERE TE_NUMTRANSAC = "' + VarToStr(GetField('TE_NUMTRANSAC')) + '" AND '
                                             + ' TE_NODOSSIER = "' + VarToStr(GetField('TE_NODOSSIER')) + '" '
                                        + 'ORDER BY TE_NODOSSIER, TE_NUMTRANSAC, TE_COMMISSION, TE_NUMEROPIECE, TE_NUMLIGNE' ) ;

      {Générations des écritures de comptabilité à partir du record (21/08/06 : Nouvelle fonction de traitement)}
      TRGenererPieceCompta( TobGen, lTobEcr );
      FreeAndNil( lTobEcr ) ;

    end;

    {FQ10223 : S'il n'y a pas de message d'erreur, on poursuit l'intégration}
    AfficheMessageErreur(Ecran.Caption, 'Certaines transactions ne pourront être intégrées en comptabilité : ');

  finally
    FiniMove;
  end;

  {Par précaution, théoriquement arrivé ici, la tob ne peut être vide}
  if TobGen.Detail.Count = 0 then begin
    HShowMessage('1;' + Ecran.Caption + ';Il n''y a aucune pièce comptable à intégrer !;W;O;O;O;', '', '');
    Exit;
  end;

  {21/08/06 : Nouvelle fonction de traitement}
  TRIntegrationPieces(TobGen, True);

  TToolbarButton97(GetControl('BCHERCHE')).Click;
end ;


initialization
  RegisterClasses ( [ TOF_TRINTEGRECOMPTA ] ) ;
end.
