{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : OUTILAFFICHELOG ()
Mots clefs ... : TOF;OUTILAFFICHELOG
*****************************************************************}
unit uTOFAfficheLog;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  dbtables,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  UTob,
  Vierge,
  Windows,
  Htb97,
  QrGrid,
  HSysMenu;


  (*
  type
  TRapport = RECORD
               Fichier : string;
               Libelle : string;

             END ;
  *)

type
  TOF_OUTILAFFICHELOG = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnChangeListe    ( Sender : Tobject );
    procedure OnChangeCBDossier( Sender : Tobject );
    procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnClickBImprimer( Sender : TObject );

  public
    Liste, CBDossier : THValComboBox;
    G: ThGrid;
    FNbErreur: THLabel;
    FTobRapport : Tob;
    FTobErreur  : Tob;

    FHSystemMenu: THSystemMenu;
  private
    function GetLibelleFromNomFichier(vCheminFichier: string): string;
    procedure InitGrille ( vCheminFichier : string );
    procedure ChargeCBDossier( vTob : Tob );
  end;

implementation

uses UControlCP;

procedure TOF_OUTILAFFICHELOG.OnNew;
begin
  inherited;
end;

procedure TOF_OUTILAFFICHELOG.OnDelete;
begin
  inherited;
end;

procedure TOF_OUTILAFFICHELOG.OnUpdate;
begin
  inherited;
end;

procedure TOF_OUTILAFFICHELOG.OnLoad;
begin
  inherited;
  Liste.ItemIndex := 0;
  Liste.OnChange(nil);
end;

procedure TOF_OUTILAFFICHELOG.OnArgument(S: string);
var
  lSt: string;
  lStParam: string;
  lLibelleRapport: string;
  lTobTemp: Tob;
  NbRapport : integer;
begin
  inherited;
  if Trim(S) = '' then
    Exit;
  lSt := S;

  FTobRapport := Tob.Create('', nil, -1);
  FTobErreur  := Tob.Create('', nil, -1);

  FNbErreur := THLabel(GetControl('FNBErreur'));
  G := THGrid(GetControl('G'));

  CBDossier := THValComboBox(GetControl('CBDOSSIER'));
  CBDossier.OnChange := OnChangeCBDossier;
  CBDossier.Clear;

  Ecran.OnKeyDown := OnKeyDownEcran;
  TToolbarButton97(GetControl('BImprimer')).OnClick := OnClickBImprimer;

  Liste := THValComboBox(GetControl('LISTE'));
  Liste.OnChange := OnChangeListe;
  Liste.Clear;

  NbRapport := 0;
  while lSt <> '' do
  begin
    lTobTemp := Tob.Create('', FTobRapport, -1);

    // Récupération du nom du fichier Journal
    lStParam := ReadTokenSt(lSt);
    lLibelleRapport := GetLibelleFromNomFichier(lStParam);
    Liste.Items.Add(lLibelleRapport);
    lTobTemp.AddChampSupValeur('LibelleRapport', lLibelleRapport, False);
    lTobTemp.AddChampSupValeur('CheminRapport', lStParam, False);
    Inc( NbRapport );
  end;

  THLabel(GetControl('FNBRapport')).Caption := IntTostr( NbRapport ) + ' rapport(s) généré(s)';
end;

procedure TOF_OUTILAFFICHELOG.OnClose;
begin
  inherited;
  if FTobRapport <> nil then FTobRapport.Free;
  if FTobErreur <> nil then  FTobErreur.Free;
end;

procedure TOF_OUTILAFFICHELOG.OnClickBImprimer(Sender: TObject);
begin
  PrintGrid([G], Ecran.Caption);
end;

procedure TOF_OUTILAFFICHELOG.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
  
    VK_Escape : TToolbarButton97(GetControl('BFerme')).Click;

    80 : if Shift=[ssCtrl] then  TToolbarButton97(GetControl('BImprimer')).Click;

  else
  end;

end;

procedure TOF_OUTILAFFICHELOG.OnChangeListe(Sender: Tobject);
var lTobRapportSelect: Tob;
begin
  if FTobRapport.Detail.Count = 0 then Exit;

  lTobRapportSelect := FTobRapport.FindFirst(['LibelleRapport'], [Liste.Items[Liste.ItemIndex]], False);
  if lTobRapportSelect <> nil then
  begin
    FTobErreur.ClearDetail;
    TobLoadFromFile(lTobRapportSelect.GetValue('CheminRapport'), nil, FTobErreur);
    G.VidePile(False);
    InitGrille(lTobRapportSelect.GetValue('CheminRapport'));
    ChargeCBDossier( FTobErreur );
    CBDossier.ItemIndex := 0;
    CBDossier.OnChange( nil );
  end;
end;

function TOF_OUTILAFFICHELOG.GetLibelleFromNomFichier(vCheminFichier: string): string;
begin
  if Pos(UpperCase(cNomRapportCpt), UpperCase(vCheminFichier)) > 0 then
    Result := TraduireMemoire('Rapport sur les comptes')

  else if Pos(UpperCase(cNomRapportMvt), UpperCase(vCheminFichier)) > 0 then
    Result := TraduireMemoire('Rapport sur les mouvements comptables')

  else if Pos(UpperCase(cNomRapportLet), UpperCase(vCheminFichier)) > 0 then
    result := TraduireMemoire('Rapport sur le lettrage')

  else if Pos(UpperCase(cNomRapportCor), UpperCase(vCheminFichier)) > 0 then
    result := TraduireMemoire('Rapport sur les corrections')

  else
    Result := TraduireMemoire('Libellé du rapport inconnu');
end;

procedure TOF_OUTILAFFICHELOG.InitGrille( vCheminFichier : string );
begin
  with G do
  begin
    Cells[0,0]   := TraduireMemoire('Dossier');
    ColWidths[0] := 55;
    ColAligns[0] := TaCenter;

    if Pos(UpperCase(cNomRapportCpt), UpperCase(vCheminFichier)) > 0 then
    begin
      ColCount := 5;
      Cells[1,0] := TraduireMemoire('Type');
      Cells[2,0] := TraduireMemoire('Compte');
      Cells[3,0] := TraduireMemoire('Libellé');
      Cells[4,0] := TraduireMemoire('Remarques');

      ColWidths[1] := 90;
      ColWidths[2] := 150;
      ColWidths[3] := 250;
      ColWidths[4] := 350;

      ColAligns[1] := TaLeftJustify;
      ColAligns[2] := TaCenter;
      ColAligns[3] := TaLeftJustify;
      ColAligns[4] := TaLeftJustify;

    end
    else
    if Pos(UpperCase(cNomRapportMvt), UpperCase(vCheminFichier)) > 0 then
    begin
      ColCount := 6;
      Cells[1,0] := TraduireMemoire('Journal');
      Cells[2,0] := TraduireMemoire('Réf. interne');
      Cells[3,0] := TraduireMemoire('Pièce/ligne');
      Cells[4,0] := TraduireMemoire('Date');
      Cells[5,0] := TraduireMemoire('Remarques');

      ColWidths[1] := 55;
      ColWidths[2] := 115;
      ColWidths[3] := 70;
      ColWidths[4] := 75;
      ColWidths[5] := 410;

      ColAligns[1] := TaCenter;
      ColAligns[3] := TaLeftJustify;
      ColAligns[3] := TaCenter;
      ColAligns[4] := TaCenter;
      ColAligns[5] := TaLeftJustify;

    end
    else
    if Pos(UpperCase(cNomRapportLet), UpperCase(vCheminFichier)) > 0 then
    begin
      ColCount := 6;
      Cells[1,0] := TraduireMemoire('Code');
      Cells[2,0] := TraduireMemoire('Auxiliaire');
      Cells[3,0] := TraduireMemoire('Compte');
      Cells[4,0] := TraduireMemoire('Etat du lettrage');
      Cells[5,0] := TraduireMemoire('Remarques');

      ColWidths[1] := 50;
      ColWidths[2] := 70;
      ColWidths[3] := 70;
      ColWidths[4] := 90;
      ColWidths[5] := 410;

      ColAligns[1] := TaCenter;
      ColAligns[2] := TaCenter;
      ColAligns[3] := TaCenter;
      ColAligns[4] := TaCenter;
      ColAligns[5] := TaLeftJustify;
    end
    else
    if Pos(UpperCase(cNomRapportCor), UpperCase(vCheminFichier)) > 0 then
    begin
      ColCount := 3;
      Cells[1,0] := TraduireMemoire('Type de traitement');
      Cells[2,0] := TraduireMemoire('Remarques');

      ColWidths[1] := 250;
      ColWidths[2] := 380;

      ColAligns[1] := TaLeftJustify;
      ColAligns[2] := TaLeftJustify;
    end;
  end;

end;

procedure TOF_OUTILAFFICHELOG.ChargeCBDossier( vTob : Tob );
var i : integer;
    lStDossier : string;
begin
  CBDossier.Clear;
  CBDossier.Items.Add('<<Tous>>');

  lStDossier := '';
  for i := 0 to vTob.Detail.Count - 1 do
  begin
    if lStDossier <> vTob.Detail[i].GetValue('Dossier') then
    begin
      lStDossier := vTob.Detail[i].GetValue('Dossier');
      CBDossier.Items.Add(lStDossier);
    end;
  end;
  
end;

procedure TOF_OUTILAFFICHELOG.OnChangeCBDossier(Sender: Tobject);
var i : integer;
    lTobErreur : Tob;
    lTobFille : Tob;
begin
  G.VidePile(False);

  if CBDossier.ItemIndex = 0 then
  begin
    FTobErreur.PutGridDetail(G, False, True, '');
  end
  else
  begin
    lTobErreur := Tob.Create('', nil, -1);
    for i := 0 to FTobErreur.Detail.Count -1 do
    begin
      if FTobErreur.Detail[i].GetValue('Dossier') = CBDossier.Text then
      begin
        lTobFille := Tob.Create('', lTobErreur, -1);
        lTobFille.Dupliquer(FTobErreur.Detail[i], False, True);
      end;
    end;
    lTobErreur.PutGridDetail(G, False, True, '');
    lTobErreur.Free;

  end;

  FNbErreur.Caption := IntToStr(G.RowCount - 1) + TraduireMemoire(' erreur(s) trouvée(s)');

  // Redimensionnement des colonnes
  TFVierge(Ecran).FormResize := True;
  FHSystemMenu.ResizeGridColumns(G);

end;


initialization
  registerclasses([TOF_OUTILAFFICHELOG]);
end.

