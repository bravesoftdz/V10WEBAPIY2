{ Unité : Source TOF de la FICHE : TRGRILLE
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 6.00.015.001   20/09/04    JP     Création de l'unité
 7.09.001.001   17/08/06    JP     Gestion du multi sociétés pour la synchronisation
 7.09.001.001   06/11/06    JP     FQ 10275 :Utilisation de la grille pour le contrôle des rubriques
                                   Ajout du Bouton d'export de la grille
 8.00.001.018   05/06/07    JP     FQ 10431 : ag_DateSynchro : ajout d'une colonne Montant
--------------------------------------------------------------------------------------}
unit TRGRILLE_TOF;

interface

uses
  Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, PrintDBG,
  {$ENDIF}
  Forms, HCtrls, HEnt1, UTOF, Vierge, UTOB, Constantes;

type
  TOF_TRGRILLE = class (TOF)
    procedure OnArgument(S : string); override;
  private
    Grille : THGrid;
    procedure DessinerGrille;
    procedure ImprimerGrille(Sender : TObject);
    procedure ExporterGrille(Sender : TObject);
    procedure MajTitre;
  end ;

procedure TRAfficheGrille(Msg : string; TypeAppel : TAfficheGrille);


implementation

uses
  HTB97, UtilPgi, Commun, HMsgBox, Dialogs, HXlsPas, SysUtils, UObjEtats,
  uTobDebug;

var
  MsgLibelle : string;
  FicheAppel : TAfficheGrille;

{---------------------------------------------------------------------------------------}
procedure TRAfficheGrille(Msg : string; TypeAppel : TAfficheGrille);
{---------------------------------------------------------------------------------------}
begin
  MsgLibelle := Msg;
  FicheAppel := TypeAppel;
  AGLLanceFiche('TR', 'TRGRILLE', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRILLE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  Grille := THGrid(GetControl('GRILLE'));
  {$IFDEF EAGLCLIENT}
  SetControlVisible('BIMPRIMER', False);
  {$ENDIF}
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := ImprimerGrille;
  TToolbarButton97(GetControl('BEXPORT'  )).OnClick := ExporterGrille;
  
  {Dessine la grille : Nb colonnes, format des cellules ...}
  DessinerGrille;
  {Met à jour le caption de la fiche et du message d'accompagnement}
  MajTitre;
  TFVierge(Ecran).LaTOF.LaTOB.PutGridDetail(Grille, True, True, '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRILLE.DessinerGrille;
{---------------------------------------------------------------------------------------}
begin
  case FicheAppel of
    ag_rien    : begin
                   Grille.ColCount := 2;
                 end;
    ag_Synchro : begin
                   {17/08/06 : une colonne de plus pour indiquer le dossier}
                   if IsTresoMultiSoc then begin
                     Grille.ColCount := 5;
                     Grille.ColWidths[0] := 50;
                     Grille.ColWidths[1] := 90;
                     Grille.ColWidths[2] := 55;
                     Grille.ColWidths[3] := 55;
                     Grille.ColWidths[4] := 270;
                     Grille.ColAligns[0] := taLeftJustify;
                     Grille.ColAligns[2] := taRightJustify;
                     Grille.ColAligns[3] := taRightJustify;
                   end else begin
                     Grille.ColCount := 4;
                     Grille.ColWidths[0] := 100;
                     Grille.ColWidths[1] := 60;
                     Grille.ColWidths[2] := 60;
                     Grille.ColWidths[3] := 300;
                     Grille.ColAligns[1] := taRightJustify;
                     Grille.ColAligns[2] := taRightJustify;
                   end;
                   Ecran.Width         := 560;
                 end;
    {06/11/06 : FQ 10275 : affichage des erreurs du paramétrage des rubrqiues (CPSUPPRUB_TOF.Pas)}
    ag_CtrlRub : begin
                   Grille.ColCount := 3;
                   Grille.ColWidths[0] := 80;
                   Grille.ColWidths[1] := 160;
                   Grille.ColWidths[2] := 300;
                   Ecran.Width         := 570;
                 end;
    {05/06/07 : FQ 10431 : Ajout du montant}
    ag_DateSynchro : begin
                       if IsTresoMultiSoc then begin
                         Grille.ColCount := 6;
                         Grille.ColWidths[0] := 45;
                         Grille.ColWidths[1] := 80;
                         Grille.ColWidths[2] := 50;
                         Grille.ColWidths[3] := 40;
                         Grille.ColWidths[4] := 70;
                         Grille.ColWidths[5] := 245;
                         Grille.ColAligns[0] := taLeftJustify;
                         Grille.ColAligns[2] := taRightJustify;
                         Grille.ColAligns[3] := taRightJustify;
                         Grille.ColAligns[4] := taRightJustify;
                       end else begin
                         Grille.ColCount := 5;
                         Grille.ColWidths[0] := 90;
                         Grille.ColWidths[1] := 60;
                         Grille.ColWidths[2] := 40;
                         Grille.ColWidths[3] := 70;
                         Grille.ColWidths[4] := 270;
                         Grille.ColAligns[1] := taRightJustify;
                         Grille.ColAligns[2] := taRightJustify;
                         Grille.ColAligns[3] := taRightJustify;
                       end;
                       Ecran.Width         := 570;
                     end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRILLE.ImprimerGrille(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TObjEtats.GenereEtatGrille(Grille, Ecran.Caption);
  {$IFDEF EAGLCLIENT}
  {$ELSE}
//  PrintDBGrid(Grille, nil, Ecran.Caption, '');
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRILLE.ExporterGrille(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  ASaveDialog : TSaveDialog;
begin
  if not ExJaiLeDroitConcept(ccExportListe, True) then Exit;
  {06/11/06 : Ajout du bouton d'export}
  ASaveDialog := TSaveDialog.Create(Ecran);
  try
    ASaveDialog.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
    ASaveDialog.DefaultExt := 'XLS';
    ASaveDialog.FilterIndex := 1;
    ASaveDialog.Options := ASaveDialog.Options + [ofOverwritePrompt, ofPathMustExist, ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];
    if ASaveDialog.Execute then
      ExportGrid(Grille, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
  finally
    FreeAndNil(ASaveDialog);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRILLE.MajTitre;
{---------------------------------------------------------------------------------------}
begin
  case FicheAppel of
    ag_rien,
    {06/11/06 : FQ 10275}
    ag_CtrlRub : begin
                   Ecran.Caption := MsgLibelle;
                   SetControlVisible('PNHAUT', False);
                 end;
    ag_Synchro : begin
                   Ecran.Caption := TraduireMemoire('Liste des écritures non synchronisées');
                   SetControlCaption('LBMESSAGE', MsgLibelle);
                 end;
    ag_DateSynchro : begin
                       Ecran.Caption := TraduireMemoire('Liste des écritures dont la date impacte les initialisations');
                       SetControlCaption('LBMESSAGE', MsgLibelle);
                     end;
  end;

  UpdateCaption(Ecran);
end;


initialization
  RegisterClasses([TOF_TRGRILLE]);

end.
