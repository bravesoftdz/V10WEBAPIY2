{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de consultation/modification de l'indemnité CP sur les
Suite ........ : mvts pris ou soldés
Mots clefs ... : PAIE;CP
*****************************************************************}
{
PT1 25/05/2004 SB V_50 FQ 11316 Pour éviter le plantage de l'exe en consultation grille non accessible
PT2 01/12/2004 PH V_60 Perte des DONNEES de la saisie manuelle de l'indemnité congés payées
               dans le cas du 1er bulletin
PT3-1 24/01/2005 SB V_60 FQ 11833 Suppression PT2
PT3-2 24/01/2005 SB V_60 Modif Caption fenètre si zoom idemnité compensatrice

}
unit UTOFIndemniteCP;

interface
uses Controls, Classes, 
{$IFNDEF EAGLCLIENT}

{$ELSE}
  eFiche, eFichlist,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, HTB97,
   Vierge, Grids, Graphics;
type
  TOF_IndemniteCP = class(TOF)
    //       procedure OnChangeField (F : TField); override ;
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
    procedure FermeWindows;
    procedure OnClose; override;
    procedure ValiderClick(Sender: TObject);
    procedure RechargeTob;
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
  private
    GGrille: THGrid;
    TobIndemnite, T_sauv: TOB;
    arg: string;
  end;

implementation

uses PgCongesPayes, P5Util;


procedure TOF_IndemniteCP.OnArgument(stArgument: string);
var
  T, TD: tob;
  Bvalider: ttoolbarbutton97;
  TypeCp : String;
begin
  inherited;
  Arg := ReadTokenSt(stArgument);
  TypeCp := ReadTokenSt(stArgument); { DEB PT3-2 }
  if TypeCP = 'SLD' then
     Ecran.Caption  := 'Indemnité compensatrice de congés payés';
  UpdateCaption (Ecran);             { FIN PT3-2 }
  BValider := Ttoolbarbutton97(getcontrol('BVALIDER'));
  if BValider <> nil then
    BValider.OnClick := Validerclick;
  GGrille := THGrid(Getcontrol('GRILLE'));
  if GGrille <> nil then
  begin
    GGrille.ColLengths[0] := -1;
    GGrille.ColLengths[1] := -1;
    GGrille.ColLengths[2] := -1;
    GGrille.ColLengths[3] := -1;
    GGrille.ColLengths[4] := -1;
    GGrille.ColLengths[5] := -1;
    if stArgument = 'Consultation' then
    begin
      GGrille.ColLengths[6] := -1;
      GGrille.Enabled := False; { PT1 }
    end;
    GGrille.ColLengths[7] := -1;
    GGrille.oncellexit := GrilleCellexit;
    GGrille.PostDrawCell := PostDrawCell;
  end;

  T_Sauv := TOB.Create('T_sauvegarde', nil, -1);

  TobIndemnite := Tob.create('tob des indemnités', nil, -1);
  T := Latob.FindFirst([''], [''], True);
  while T <> nil do
  begin
    if T.getvalue('PCN_CODERGRPT') <> -1 then
    begin
      TD := TOB.Create('ABSENCESALARIE', T_Sauv, -1);
      TD.Dupliquer(T, FALSE, TRUE, False);
      T.ChangeParent(TobIndemnite, -1);
    end;
    T := Latob.Findnext([''], [''], True);
  end;

end;

procedure TOF_IndemniteCP.PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ((GGrille.ColLengths[Acol] = -1) and (ARow > 0)) then
    GridGriseCell(GGrille, Acol, Arow, Canvas);
end;

procedure TOF_IndemniteCP.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Acol = GGrille.colcount - 2 then
    if not Isnumeric(GGrille.cells[ACol, ARow]) then
      GGrille.cells[ACol, ARow] := '';
end;

procedure TOF_IndemniteCP.FermeWindows;
var
  i: integer;
  Init: word;
  T: Tob;
begin
  if Arg = 'Consultation' then exit;
  Lasterror := 0;
  TFVierge(Ecran).Retour := '0';
  //N°Titre;Texte;Icône; Boutons ; Bouton Défaut ; Bouton Cancel , FURTIF, BEEP, HelpID
  // PT2 Ergonomie sur le contenu du message 
  Init := HShowMessage('1;Sauvegarde et fermeture;Voulez-vous enregistrer votre saisie?;N;YNC;Y;', '', '');
  if Init = mrYes then // je réinitialise la grille
  begin // B3
    TobIndemnite.GetGridDetail(GGrille, -1, '', 'PCN_LIBELLE;PCN_MVTDUPLIQUE;PCN_JOURS' +
      ';PCN_VALOX;PCN_VALOMS;PCN_VALORETENUE;PCN_VALOMANUELLE;PCN_PERIODEPAIE');
    T := TobIndemnite.findfirst([''], [''], True);
    for i := 1 to GGrille.rowcount - 1 do
    begin
      if not IsValeurnumeric(GGrille.cells[GGrille.colcount - 2, i]) then
        GGrille.cells[GGrille.colcount - 2, i] := '';
    end;

    i := 1;
    while T <> nil do
    begin
      if GGrille.cells[GGrille.colcount - 2, i] <> '' then
      begin // PT2  Rajout Initialisation de la date de fin pour provoquer le rechargement des CP pris
        T.putvalue('PCN_MODIFVALO', 'X');
        //if FirstBull then T.putvalue('PCN_DATEFIN', IDate1900); PT3-1
      end
      else
      begin
        T.putvalue('PCN_MODIFVALO', '-');
        T.Putvalue('PCN_VALOMANUELLE', 0);
      end;
      T := TobIndemnite.findnext([''], [''], True);
      i := i + 1;
    end;

    //if FirstBull then TobIndemnite.InsertOrUpdateDB(FALSE); // PT2 Ecriture dans la base uniquement dans le cas du 1er bulletin PT3-1
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

procedure TOF_IndemniteCP.RechargeTob;
var
  T: Tob;
begin
  T := TobIndemnite.FindFirst([''], [''], True);
  while T <> nil do
  begin
    T.ChangeParent(Latob, -1);
    T := TobIndemnite.Findnext([''], [''], True);
  end;
end;

procedure TOF_IndemniteCP.OnLoad;
var
  t: tob;
  i: integer;
begin
  if GGrille <> nil then
    TobIndemnite.PutGridDetail(GGrille, false, false,
      'PCN_LIBELLE;PCN_MVTDUPLIQUE;PCN_JOURS;PCN_VALOX;PCN_VALOMS;' +
      'PCN_VALORETENUE;PCN_VALOMANUELLE;PCN_PERIODEPAIE'
      , false);
  T := TobIndemnite.findfirst([''], [''], True);
  i := 1;
  while T <> nil do
  begin
    if T.getvalue('PCN_MODIFVALO') <> 'X' then
      GGrille.cells[GGrille.colcount - 2, i] := '';
    T := TobIndemnite.findnext([''], [''], True);
    i := i + 1;
  end;
end;
// Fonction de validation de la saisie

procedure TOF_IndemniteCP.ValiderClick(Sender: TObject);
begin
  //OnCLose;
end;

procedure TOF_IndemniteCP.OnClose;
begin
  FermeWindows;
end;
initialization
  registerclasses([TOF_IndemniteCP]);

end.

