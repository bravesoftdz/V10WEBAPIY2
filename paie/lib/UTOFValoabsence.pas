{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion de la valorisation des absences
Mots clefs ... : PAIE;CP
*****************************************************************}
{
PT1 01/12/2004 PH V_60 Perte des DONNEES de la saisie manuelle de l'absence congés payées
               dans le cas du 1er bulletin
PT2 27/01/2005 SB V_60 FQ 11833 Suppression PT2
}

unit UTOFValoAbsence;

interface
uses Controls, Classes,
{$IFNDEF EAGLCLIENT}
{$ELSE}
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOB, UTOF, HTB97,
  Vierge, Graphics, Grids;

type
  TOF_ValoAbsence = class(TOF)
    //       procedure OnChangeField (F : TField); override ;
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure FermeWindows;
    procedure OnClose; override;
    procedure ValiderClick(Sender: TObject);
    procedure RechargeTob;
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
  private
    GGrille: THGrid;
    TobAbsence, T_sauv: TOB;
  end;

implementation

uses P5Util;

procedure TOF_ValoAbsence.OnArgument(stArgument: string);
var
  T, TD: tob;
  BValider: Ttoolbarbutton97;
begin
  inherited;

  GGrille := THGrid(Getcontrol('GRILLE'));
  if GGrille <> nil then
  begin
    GGrille.ColLengths[0] := -1;
    GGrille.ColLengths[1] := -1;
    GGrille.ColLengths[2] := -1;
    GGrille.ColLengths[3] := -1;
    if stArgument = 'Consultation' then GGrille.ColLengths[4] := -1; //SB 04/07/2001 gestion du mode consultation
    GGrille.oncellexit := GrilleCellexit;
    GGrille.PostDrawCell := PostDrawCell;
  end;

  T_Sauv := TOB.Create('T_sauvegarde', nil, -1);
  BValider := Ttoolbarbutton97(getcontrol('BVALIDER'));
  if BValider <> nil then
    BValider.OnClick := Validerclick;

  TobAbsence := Tob.create('tob des absences', nil, -1);
  T := Latob.FindFirst([''], [''], True);
  while T <> nil do
  begin

    if T.getvalue('PCN_CODERGRPT') <> -1 then
    begin
      TD := TOB.Create('ABSENCESALARIE', T_Sauv, -1);
      TD.Dupliquer(T, FALSE, TRUE, False);
      T.ChangeParent(TobAbsence, -1);
    end;
    T := Latob.Findnext([''], [''], True);
  end;

end;

procedure TOF_ValoAbsence.PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ((GGrille.ColLengths[Acol] = -1) and (ARow > 0)) then
    GridGriseCell(GGrille, Acol, Arow, Canvas);
end;

procedure TOF_ValoAbsence.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Acol = GGrille.colcount - 1 then
    if not Isnumeric(GGrille.cells[ACol, ARow]) then
      GGrille.cells[ACol, ARow] := '';
end;

procedure TOF_ValoAbsence.ValiderClick(Sender: TObject);
begin

end;

procedure TOF_ValoAbsence.FermeWindows;
var
  Init: word;
  T: Tob;
  i: integer;
begin

  Lasterror := 0;
  TFVierge(Ecran).Retour := '0';
  //N°Titre;Texte;Icône; Boutons ; Bouton Défaut ; Bouton Cancel , FURTIF, BEEP, HelpID
  Init := HShowMessage('1;Sauvegarde et fermeture;Voulez-vous enregistrer votre saisie ?;N;YNC;Y;', '', '');// PT1 Ergonomie
  if Init = mrYes then // je réinitialise la grille
  begin // B3
    for i := 1 to GGrille.rowcount - 1 do
    begin
      if not Isnumeric(GGrille.cells[GGrille.colcount - 1, i]) then
        GGrille.cells[GGrille.colcount - 1, i] := '';
    end;
    TobAbsence.GetGridDetail(GGrille, -1, '', 'PCN_LIBELLE;PCN_MVTDUPLIQUE;PCN_JOURS;PCN_ABSENCE;PCN_ABSENCEMANU');
    T := TobAbsence.findfirst([''], [''], True);
    i := 1;
    while T <> nil do
    begin
      if GGrille.cells[GGrille.colcount - 1, i] <> '' then
      begin
        T.putvalue('PCN_MODIFABSENCE', 'X');
// PT2  if FirstBull then T.putvalue('PCN_DATEFIN', IDate1900);  // PT1  Rajout Initialisation de la date de fin pour provoquer le rechargement des CP pris
      end
      else
      begin
        T.putvalue('PCN_MODIFABSENCE', '-');
        T.Putvalue('PCN_ABSENCEMANU', 0);
      end;

      T := TobAbsence.findnext([''], [''], True);
      i := i + 1;
    end;
//PT2  if FirstBull then TobAbsence.InsertOrUpdateDB(FALSE); // PT1 Ecriture dans la base uniquement dans le cas du 1er bulletin
    RechargeTob;
  end else
    if Init = mrNo then
  begin
    t := T_Sauv.findfirst([''], [''], True);
    while t <> nil do
    begin
      T.ChangeParent(LaTob, -1);
      T := T_Sauv.Findnext([''], [''], True);
    end;
    if T_Sauv.findfirst([''], [''], True) <> nil then T_sauv.free;
    TFVierge(Ecran).Retour := '-1';
    exit;
  end
  else
  begin
    LastErrorMsg := '';
    LastError := 1;
  end;

end;

procedure TOF_ValoAbsence.RechargeTob;
var
  T: Tob;
begin
  T := TobAbsence.FindFirst([''], [''], True);
  while T <> nil do
  begin
    T.ChangeParent(Latob, -1);
    T := TobAbsence.Findnext([''], [''], True);
  end;
end;

procedure TOF_ValoAbsence.OnLoad;
var
  i: integer;
  T: tob;

begin
  if GGrille <> nil then
    TobAbsence.PutGridDetail(GGrille, false, false,
      'PCN_LIBELLE;PCN_MVTDUPLIQUE;PCN_JOURS;PCN_ABSENCE;PCN_ABSENCEMANU', false);
  T := TobAbsence.findfirst([''], [''], True);
  i := 1;
  while T <> nil do
  begin
    if T.getvalue('PCN_MODIFABSENCE') <> 'X' then
      GGrille.cells[GGrille.colcount - 1, i] := '';
    T := TobAbsence.findnext([''], [''], True);
    i := i + 1;
  end;
end;


procedure TOF_ValoAbsence.OnUpdate;
begin

end;

procedure TOF_ValoAbsence.OnClose;
begin
  FermeWindows;
end;
initialization
  registerclasses([TOF_ValoAbsence]);

end.

