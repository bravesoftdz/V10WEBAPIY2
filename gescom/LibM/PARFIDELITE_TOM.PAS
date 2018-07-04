{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Source TOM des TABLES :
Suite ........ :   PARFIDELITE (Paramètrage de la fidélité client)
Suite ........ :   PARREGLEFID (Paramétrage des règles de fidélité)
Mots clefs ... : TOM;PARFIDELITE;PARREGLEFID
*****************************************************************}
unit PARFIDELITE_TOM;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eFiche, eFichList,
  {$ELSE}
  db, dbtables, Fiche, FichList,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOM, UTob,
  Windows, Messages, Spin, uTableFiltre, SaisieList, M3FP, SaisUtil, UtilArticle,
  FactArticle, HTB97;

type
  // TOM de la TABLE : PARFIDELITE
  TOM_PARFIDELITE = class(TOM)
    GEtab, GEtabSelect : THGrid;
    TOBEtab, TOBEtabSelect : TOB;
    BFlecheGauche, BFlecheDroite, BFlecheTous, BFlecheAucun : TToolBarButton97;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
    procedure BFlecheGaucheClick(Sender: TObject);
    procedure BFlecheDroiteClick(Sender: TObject);
    procedure BFlecheTousClick(Sender: TObject);
    procedure BFlecheAucunClick(Sender: TObject);
  private
    lMono: boolean; // lancement MONOFICHE
    procedure ActiveChamp(NomChamp: string; Active: boolean);
    procedure OnChangeTypeDureeFid;
    //procedure OnChangeRemiseFidelite;
    procedure SetLastError(Num: integer; ou: string);
    procedure UtilBonAchatOnClick;
    procedure ChargementEtablissement;
    procedure DroiteVersGauche(Ligne : Integer);
    procedure GaucheVersDroite(Ligne : integer);
    function AutoriseFleche : Boolean;
  end;

  // Source TOM de la TABLE : PARREGLEFID
  TOM_PARREGLEFID = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    TF: TTableFiltre;
    TOBSeuil: TOB; // Liste des seuils d'une règle de fidélité
    GS: THGrid; // Grille de saisie des seuils
    SeuilModifie: boolean; // Indicateur de modification des seuils
    StCellCur: string; // Valeur de la cellule de la grille des seuils avant modification
    procedure ChargementSeuil;
    procedure GSEnter(Sender: TObject);
    procedure GSExit(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GSElipsisClick(Sender: TObject);
    procedure InsertLigneGS;
    procedure AffichageRegle(Regle: string);
    procedure FormateGrid(Regle: string);
    procedure SupLigneClick(Sender: TObject);
  end;

const
  TexteMessage: array[1..3] of string = (
    {1}'Vous devez renseigner une option de démarrage'
    {2}, 'Vous devez renseigner au minimum une nature de pièce'
    {3}, 'Vous devez renseigner au minimum un type d''article'
    );

  Champs_GS = 'GFS_SEUILMINI;GFS_SEUILMAXI;GFS_CONSTANTE;GFS_COEF;GFS_ARRONDI;GFS_VALEURMAXIFID;GFS_CODEARTICLE;GFS_QTEFACT';

implementation

{*****************************************************************
**  PARFIDELITE (Paramètrage de la fidélité client)             **
*****************************************************************}

procedure TOM_PARFIDELITE.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TraduireMemoire(TexteMessage[LastError]);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Active ou désactive un champ et son libellé
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.ActiveChamp(NomChamp: string; Active: boolean);
begin
  SetControlEnabled(NomChamp, Active);
  SetControlEnabled('T' + NomChamp, Active);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Changement de GFO_TYPEDUREEFID : Type de durée
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnChangeTypeDureeFid;
var
  Stg: string;
  Ok: boolean;
  Ctrl: TControl;
begin
  Stg := GetField('GFO_TYPEDUREEFID');
  Ok := (Stg = '002');
  ActiveChamp('GFO_DATEFIN', Ok);
  if not Ok then
    SetField('GFO_DATEFIN', idate2099);

  Ok := (Stg = '003');
  ActiveChamp('GFO_NBJOUR', Ok);
  Ctrl := GetControl('GFO_NBJOUR');
  if (Ctrl <> nil) and (Ctrl is TSpinEdit) then
  begin
    with TSpinEdit(Ctrl) do
    begin
      if Ok then
      begin
        MinValue := 1;
        MaxValue := 999;
        MaxLength := 3;
        if (Value < 1) or (Value > 999) then
          Value := 365;
      end else
      begin
        MinValue := 0;
        MaxValue := 0;
        MaxLength := 0;
        Value := 0;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Changement de GFO_REMISEFIDELITE : Remise de
Suite ........ : fidélité
Mots clefs ... :
*****************************************************************}
{
procedure TOM_PARFIDELITE.OnChangeRemiseFidelite;
var
  Ok: boolean;
begin
  Ok := (GetField('GFO_REMISEFIDELITE') = 'X');
  ActiveChamp('GFO_TYPEREMISE', Ok);
  if not Ok then
    SetField('GFO_TYPEREMISE', '');

  ActiveChamp('GFO_MODEPAIE', (not Ok));
  if Ok then
    SetField('GFO_MODEPAIE', '');
end; }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnNewRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnNewRecord;
begin
  inherited;
  SetField('GFO_TYPEDUREEFID', '001');
  SetField('GFO_TYPECUMULFID', '101');
  SetField('GFO_NATUREPIECEFID', 'FFO;');
  SetField('GFO_TYPEARTICLEFID', 'MAR;');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnDeleteRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnDeleteRecord;
var
  sSql: string;
begin
  inherited;
  sSql := 'SELECT GFR_CODEFIDELITE FROM PARREGLEFID'
    + ' WHERE GFR_CODEFIDELITE="' + GetField('GFO_CODEFIDELITE') + '"';
  if ExisteSQL(sSql) then
  begin
    LastError := 1;
    LastErrorMsg := 'Ce programme de fidélité contient encore des règles, son annulation est impossible.';
    Exit;
  end;

  sSql := 'SELECT ET_PROGRAMME FROM ETABLISS'
    + ' WHERE ET_PROGRAMME="' + GetField('GFO_CODEFIDELITE') + '"';
  if ExisteSQL(sSql) then
  begin
    LastError := 1;
    LastErrorMsg := 'Ce programme de fidélité est utilisé par un établissement, son annulation est impossible.';
    Exit;
  end;

  if (Ecran <> nil) and (Ecran is TFFiche) then
  begin
    TFFiche(Ecran).Retour := GetField('GFO_CODEFIDELITE');
    if lMono then
      PostMessage(TWinControl(Ecran).Handle, WM_CLOSE, 0, 0);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnUpdateRecord;
var i : integer;
    CodProg : String;
begin
  inherited;
  if (GetField('GFO_DEMARCLIENT') = '-') and (GetField('GFO_DEMARDEMANDE') = '-')
    and (GetField('GFO_DEMARREGLE') = '-') then SetLastError(1, 'GFO_DEMARCLIENT');

  if GetField('GFO_NATUREPIECEFID') = '' then SetLastError(2, 'GFO_NATUREPIECEFID');

  if GetField('GFO_TYPEARTICLEFID') = '' then SetLastError(3, 'GFO_TYPEARTICLEFID');

  if TFFiche(Ecran).FTypeAction = taModif then
  begin
    For i:=0 to TOBEtab.Detail.Count-1 do
    begin
      TOBEtab.Detail[i].PutValue('ET_PROGRAMME', '');
      TOBEtab.Detail[i].PutValue('ET_UTILISATEUR', V_PGI.User);
    end;
    CodProg := GetControlText('GFO_CODEFIDELITE');
    For i:=0 to TOBEtabSelect.Detail.Count-1 do
    begin
      TOBEtabSelect.Detail[i].PutValue('ET_PROGRAMME', CodProg);
      TOBEtabSelect.Detail[i].PutValue('ET_UTILISATEUR', V_PGI.User);
    end;
    TOBEtab.UpdateDB();
    TOBEtabSelect.UpdateDB();
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnAfterUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnAfterUpdateRecord;
begin
  inherited;
  if (Ecran <> nil) and (Ecran is TFFiche) then
  begin
    TFFiche(Ecran).Retour := GetField('GFO_CODEFIDELITE');
    if lMono then
      PostMessage(TWinControl(Ecran).Handle, WM_CLOSE, 0, 0);
  end;
end;

procedure TOM_PARFIDELITE.ChargementEtablissement;
var StSQL, CodeFid : String;
    i : integer;
begin
  CodeFid := GetControlText('GFO_CODEFIDELITE');
  StSQL := 'Select ET_ETABLISSEMENT, ET_LIBELLE, ET_PROGRAMME From ETABLISS ' +
  'Where ET_PROGRAMME = "" OR ET_PROGRAMME = "' + CodeFid + '"';
  if TOBEtab<>Nil then begin TOBEtab.Free; TOBEtab := Nil; end;
  if TOBEtabSelect<>Nil then begin TOBEtabSelect.Free; TOBEtabSelect := Nil; end;
  TOBEtab := TOB.Create('Liste etab', Nil, -1);
  TOBEtabSelect := TOB.Create('Liste etab', Nil, -1);
  TOBEtab.LoadDetailDBFromSQL('ETABLISS', StSQL);

  if TFFiche(Ecran).FTypeAction <> taCreat then
  begin
    For i:=TOBEtab.Detail.Count-1 Downto 0 do
    begin
      if TOBEtab.Detail[i].GetValue('ET_PROGRAMME') = CodeFid then
        TOBEtab.Detail[i].ChangeParent(TOBEtabSelect, -1);
    end;
  end;

  GEtab.VidePile(False);
  GEtabSelect.VidePile(False);

  TOBEtab.PutGridDetail(GEtab, False, False, 'ET_ETABLISSEMENT;ET_LIBELLE');
  TOBEtabSelect.PutGridDetail(GEtabSelect, False, False, 'ET_ETABLISSEMENT;ET_LIBELLE');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnLoadRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnLoadRecord;
begin
  inherited;
  UtilBonAchatOnClick;
  ChargementEtablissement;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnChangeField
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnChangeField(F: TField);
begin
  inherited;
  if F.FieldName = 'GFO_TYPEDUREEFID' then
    OnChangeTypeDureeFid
      //else if F.FieldName = 'GFO_REMISEFIDELITE' then
    //  OnChangeRemiseFidelite
  else ;
end;

procedure TOM_PARFIDELITE.DroiteVersGauche(Ligne : Integer);
var Nb : integer;
    TOBEtabSelection : TOB;
    Etab : string;
begin
  // Changement du parent
  TOBEtabSelection := TOBEtabSelect.FindFirst(['ET_ETABLISSEMENT'],[GEtabSelect.Cells[0,Ligne]], False);
  if TOBEtabSelection = Nil then Exit;
  if TOBEtabSelection.GetValue('ET_PROGRAMME') <> '' then
  begin
    Etab := TOBEtabSelection.GetValue('ET_ETABLISSEMENT');
    if ExisteSQL('SELECT GFE_TIERS FROM FIDELITEENT WHERE GFE_ETABLISSEMENT="' +
    Etab + '" AND GFE_PROGRAMME="' + TOBEtabSelection.GetValue('ET_PROGRAMME') + '"') then
    begin
      PGIInfo( TraduireMemoire('Impossible d''enlever l''établissement (')
      + RechDom('TTETABLISSEMENT', Etab, False) +
      TraduireMemoire(') car des cartes de fidélité ont déjà été créées avec ce programme.') );
      Exit;
    end;
  end;
  TOBEtabSelection.ChangeParent(TOBEtab, -1);
  // Supprime la ligne du tableau
  GEtabSelect.DeleteRow(Ligne);
  Nb := TOBEtab.Detail.Count;
  // Augmente le nombre de lignes du tableau qui reçoit
  GEtab.RowCount := Nb + 1;
  // Affiche l'établissement enlevé
  TOBEtab.Detail[Nb-1].PutLigneGrid(GEtab, Nb ,False, False, 'ET_ETABLISSEMENT;ET_LIBELLE');
end;

function TOM_PARFIDELITE.AutoriseFleche : Boolean;
begin
  Result := False;
  if TFFiche(Ecran).FTypeAction = taCreat then
  begin
    PGIInfo(TraduireMemoire('Action impossible en création'));
    Exit;
  end
  else if TFFiche(Ecran).FTypeAction = taConsult then Exit;
  TFFiche(Ecran).QFiche.Edit;
  Result := True;
end;

procedure TOM_PARFIDELITE.BFlecheGaucheClick(Sender: TObject);
begin
  // Y a t il quelque chose de sélectionné ?
  if GEtabSelect.Row < 0 then exit;
  if not AutoriseFleche then Exit;
  DroiteVersGauche(GEtabSelect.Row);
end;

procedure TOM_PARFIDELITE.GaucheVersDroite(Ligne : integer);
var Nb : integer;
    TOBEtabSelection : TOB;
begin
  // Changement du parent de l'élément de la liste des établissements
  TOBEtabSelection := TOBEtab.FindFirst(['ET_ETABLISSEMENT'],[GEtab.Cells[0,Ligne]], False);
  if TOBEtabSelection = Nil then Exit;
  TOBEtabSelection.ChangeParent(TOBEtabSelect, -1);
  // Supprime la ligne du tableau
  GEtab.DeleteRow(Ligne);
  Nb := TOBEtabSelect.Detail.Count;
  // Augmente le nombre de lignes du tableau qui reçoit
  GEtabSelect.RowCount := Nb + 1;
  // Affiche l'établissement ajouté
  TOBEtabSelect.Detail[Nb-1].PutLigneGrid(GEtabSelect, Nb ,False, False, 'ET_ETABLISSEMENT;ET_LIBELLE');
end;

procedure TOM_PARFIDELITE.BFlecheDroiteClick(Sender: TObject);
begin
  // Y a t il quelque chose de sélectionné ?
  if GEtab.Row < 0 then exit;
  if not AutoriseFleche then Exit;
  GaucheVersDroite(GEtab.Row);
end;

procedure TOM_PARFIDELITE.BFlecheTousClick(Sender: TObject);
var i : integer;
begin
  if not AutoriseFleche then Exit;
  For i:=TOBEtab.Detail.Count-1 Downto 0 do GaucheVersDroite(i+1);
end;

procedure TOM_PARFIDELITE.BFlecheAucunClick(Sender: TObject);
var i : integer;
begin
  if not AutoriseFleche then Exit;
  For i:=TOBEtabSelect.Detail.Count-1 Downto 0 do DroiteVersGauche(i+1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnArgument(S: string);
begin
  lMono := (pos('MONOFICHE', S) > 0);
  inherited;

  GEtab := THGrid(GetControl('GETAB'));
  GEtabSelect := THGrid(GetControl('GETABSELECT'));

  BFlecheGauche := TToolbarButton97(GetControl('BFLECHEGAUCHE'));
  BFlecheDroite := TToolbarButton97(GetControl('BFLECHEDROITE'));
  BFlecheAucun := TToolbarButton97(GetControl('BFLECHEAUCUN'));
  BFlecheTous := TToolbarButton97(GetControl('BFLECHETOUS'));

  if (TFFiche(Ecran).FTypeAction = taConsult) then
  begin
    BFlecheGauche.Enabled := False;
    BFlecheDroite.Enabled := False;
    BFlecheAucun.Enabled := False;
    BFlecheTous.Enabled := False;
  end
  else
  begin
    BFlecheGauche.OnClick := BFlecheGaucheClick;
    BFlecheDroite.OnClick := BFlecheDroiteClick;
    BFlecheAucun.OnClick := BFlecheAucunClick;
    BFlecheTous.OnClick := BFlecheTousClick;
  end;

  GEtab.OnDblClick := BFlecheDroiteClick;
  GEtabSelect.OnDblClick := BFlecheGaucheClick;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnClose;
begin
  inherited;
  if TOBEtab<>Nil then begin TOBEtab.Free; TOBEtab := Nil; end;
  if TOBEtabSelect<>Nil then begin TOBEtabSelect.Free; TOBEtabSelect := Nil; end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnCancelRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARFIDELITE.OnCancelRecord;
begin
  inherited;
  ChargementEtablissement;
end;

{*****************************************************************
**  PARREGLEFID (Paramétrage des règles de fidélité)            **
*****************************************************************}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnNewRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnNewRecord;
begin
  inherited;
  SetField('GFR_DATEFINPERIOD', iDate2099);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnDeleteRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnDeleteRecord;
var
  sSql: string;
begin
  inherited;
  sSql := 'DELETE PARSEUILFID WHERE GFS_CODEFIDELITE="' + GetField('GFR_CODEFIDELITE') + '"'
    + ' AND GFS_REGLEFIDELITE="' + GetField('GFR_REGLEFIDELITE') + '"';
  ExecuteSQL(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnUpdateRecord;
var
  sSql, sProg, sRegle: string;
  Ind, NbR: integer;
  TOBL: TOB;
  NowFutur: TDateTime;
begin
  inherited;
  if TFSaisieList(Ecran).FTypeAction = taConsult then Exit;

  if GetControlText('GFR_REGLEFIDELITE') = '' then
  begin
    SetFocusControl('GFR_REGLEFIDELITE');
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Vous devez renseigner un code.');
  end;

  // Traitement des seuils de la règle de fidélité
  if SeuilModifie then
  begin
    // Lecture de la grille avec suppression des lignes vides
    NbR := 0;
    for Ind := GS.RowCount - 1 downto GS.FixedRows do
      if GS.Cells[0, Ind] <> '' then
      begin
        NbR := (Ind + 1 - GS.FixedRows);
        Break;
      end
      else GS.DeleteRow(Ind);

    // Affecte le seuil maxi
    for Ind := GS.RowCount - 1 downto GS.FixedRows do
      if (Ind = GS.RowCount - 1) then GS.Cells[1, Ind] := '99999999'
      else GS.Cells[1, Ind] := GS.Cells[0, Ind + 1];

    if NbR > 0 then TOBSeuil.GetGridDetail(GS, NbR, 'PARSEUILFID', Champs_GS);

    sProg := GetField('GFR_CODEFIDELITE');
    sRegle := GetField('GFR_REGLEFIDELITE');
    // suppression des anciens seuils
    sSql := 'DELETE PARSEUILFID WHERE GFS_CODEFIDELITE="' + sProg + '"'
      + ' AND GFS_REGLEFIDELITE="' + sRegle + '"';
    ExecuteSQL(sSql);

    if (TOBSeuil.Detail.Count > 0) and (V_PGI.IoError = oeOk) then
    begin
      // renumérotation des seuils
      NowFutur := NowH;
      for Ind := 0 to TOBSeuil.Detail.Count - 1 do
      begin
        TOBL := TOBSeuil.Detail[Ind];
        if TOBL.Modifie then
        begin
          TOBL.SetDateModif(NowFutur);
          TOBL.UpdateDateModif;
        end;
        TOBL.PutValue('GFS_CODEFIDELITE', sProg);
        TOBL.PutValue('GFS_REGLEFIDELITE', sRegle);
        TOBL.PutValue('GFS_RANG', (Ind + 1));
      end;
      // insertion des nouveaux seuils
      TOBSeuil.SetAllModifie(True);
      if not TOBSeuil.InsertDB(nil) then V_PGI.IoError := oeUnknown;
    end;
  end;

  if (V_PGI.IoError = oeOk) and (LastError = 0) then
    ExecuteSQL('UPDATE PARFIDELITE SET GFO_DATEMODIF = "' + USTime(NowH) + '", ' +
    'GFO_UTILISATEUR = "' + V_PGI.User + '" WHERE GFO_CODEFIDELITE = "' + sProg + '"');

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnAfterUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnAfterUpdateRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnLoadRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnLoadRecord;
begin
  inherited;
  // chargement des seuils de la règle de fidélité
  ChargementSeuil;
end;

procedure TOM_PARREGLEFID.AffichageRegle(Regle: string);
var StVal: string;
    OldChange : TNotifyEvent;
    C : THValComboBox;
begin
  if Regle='' then Exit;
  if (Regle='001') or (Regle='002') or (Regle='003') then
  begin
    PGIInfo(TraduireMemoire('Cette règle n''est pas traitée dans cette version'));
    Exit;
  end;

  if GetField('GFR_LIBELLE') = '' then
    SetField('GFR_LIBELLE', RechDom('GCTYPEREGLEFID', GetControlText('GFR_TYPEREGLEFID'), False));
  SetControlProperty('Page4', 'TabVisible', Regle <> '001');

  OldChange := nil;
  C := THValComboBox( GetControl('GFR_CHAMPSEUIL') );
  if  Assigned( C.OnChange ) then
  begin
    OldChange := C.OnChange;
    C.OnChange := nil;
  end;

  StVal := GetField('GFR_CHAMPSEUIL');
  if Regle = '005' then SetControlProperty('GFR_CHAMPSEUIL', 'Plus', 'AND CO_LIBRE LIKE "%L%"')
  else SetControlProperty('GFR_CHAMPSEUIL', 'Plus', 'AND CO_LIBRE LIKE "%C%"');
  SetControlText('GFR_CHAMPSEUIL', StVal);

  if Assigned( OldChange ) then C.OnChange := OldChange;

  SetControlVisible('GFR_CHAMPSEUIL', Regle <> '006');
  SetControlVisible('TGFR_CHAMPSEUIL', Regle <> '006');
  SetControlVisible('GFR_NBJOUR', Regle = '006');
  SetControlVisible('TGFR_NBJOUR', Regle = '006');

  FormateGrid(Regle);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnChangeField
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnChangeField(F: TField);
begin
  inherited;
  if F.FieldName = 'GFR_TYPEREGLEFID' then AffichageRegle(GetField('GFR_TYPEREGLEFID'));
  //else if F.FieldName = 'GFR_CHAMPSEUIL' then SetControlText('GFR_CHAMPSEUIL', GetField('GFR_CHAMPSEUIL'));
end;

procedure TOM_PARREGLEFID.SupLigneClick(Sender: TObject);
Var ARow : Integer;
begin
  ARow := GS.Row;
  if ARow <= 1 then Exit;
  GS.DeleteRow(ARow);
  //GS.Row := ARow - 1;
  //SendMessage(GS.Handle, WM_KEYUP, VK_UP, 0);
  SeuilModifie := True;
  if TF.State = dsBrowse then TF.State := dsEdit;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnArgument(S: string);
var
  Ctrl: TControl;
begin
  inherited;
  if (Ecran <> nil) and (Ecran is TFSaisieList) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    TOBSeuil := TOB.Create('', nil, -1);
    Ctrl := GetControl('GS');
    if (Ctrl <> nil) and (Ctrl is THGrid) then
    begin
      GS := THGrid(Ctrl);
      GS.OnEnter := GSEnter;
      GS.OnExit := GSExit;
      GS.OnCellEnter := GSCellEnter;
      GS.OnCellExit := GSCellExit;
      GS.OnKeyDown := GSKeyDown;
      GS.OnElipsisClick := GSElipsisClick;
    end;
    TToolBarButton97(GetControl('BDELLIGNE')).OnClick := SupLigneClick;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnClose;
begin
  inherited;
  if TOBSeuil <> nil then FreeAndNil(TOBSeuil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnCancelRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.OnCancelRecord;
begin
  inherited;
  // chargement des seuils de la règle de fidélité
  ChargementSeuil;
end;

procedure TOM_PARREGLEFID.FormateGrid(Regle: string);
var iCol : integer;
begin
  with GS do
  begin
    for iCol := 0 to ColCount - 1 do ColLengths[iCol] := 0;
    ColWidths[0] := 50;
    ColAligns[0] := taRightJustify;
    ColTypes[0] := 'R'; //A partir de
    ColWidths[1] := -1;
    ColLengths[1] := -1; //Seuil maxi
    ColWidths[2] := 40;
    ColAligns[2] := taRightJustify;
    ColTypes[2] := 'R'; //Constante
    ColWidths[3] := 30;
    ColAligns[3] := taRightJustify;
    ColTypes[3] := 'R'; //Coeff
    ColWidths[5] := 60;
    ColAligns[5] := taRightJustify;
    ColTypes[5] := 'R'; //Valeur Maxi
    ColWidths[6] := 60; //Code Article
    ColWidths[7] := 40;
    ColAligns[7] := taRightJustify;
    ColTypes[7] := 'R'; //Qte fact
    if (Regle = '002') or (Regle = '006') then
    begin
      ColWidths[3] := -1; ColLengths[3] := -1;
      ColWidths[4] := -1; ColLengths[4] := -1;  //gs.ValCombo := Nil;
      ColWidths[5] := -1; ColLengths[5] := -1;
      GS.ColFormats[4] := '';
    end
    else
    begin
      ColWidths[4] := 60;
      GS.ColFormats[4] := 'CB=GCCODEARRONDI';
      ColWidths[6] := -1; ColLengths[6] := -1;
      ColWidths[7] := -1; ColLengths[7] := -1;
    end;
  end;
  GS.Col := 0; GS.Row := 1;
  AffecteGrid(GS, TFSaisieList(Ecran).FTypeAction);
  TFSaisieList(Ecran).HMTrad.ResizeGridColumns(GS);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Chargement des seuils de la règle de fidélité
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.ChargementSeuil;
var Stg: string;
begin
  SeuilModifie := False;
  if GS.valcombo <> Nil then GS.valcombo.visible := false;
  GS.ElipsisButton := False;
  if TOBSeuil = nil then Exit;

  Stg := '"' + GetField('GFR_CODEFIDELITE') + '";"' + GetField('GFR_REGLEFIDELITE') + '"';
  TOBSeuil.LoadDetailDB('PARSEUILFID', Stg, 'GFS_RANG', nil, False);

  GS.VidePile(False);

  FormateGrid(GetField('GFR_TYPEREGLEFID'));

  TOBSeuil.PutGridDetail(GS, False, False, Champs_GS, True);

  StCellCur := GS.Cells[GS.Col, GS.Row];

  TFSaisieList(Ecran).HMTrad.ResizeGridColumns(GS);
end;


procedure TOM_PARREGLEFID.GSExit(Sender: TObject);
begin
  SetControlVisible('BDELLIGNE',False);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : GSEnter
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.GSEnter(Sender: TObject);
var
  Cancel: boolean;
  ACol, ARow: integer;
begin
  Cancel := False;
  ACol := GS.Col;
  ARow := GS.Row;
  GSCellEnter(GS, ACol, ARow, Cancel);
  SetControlVisible('BDELLIGNE',True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : GSCellEnter
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  StCellCur := GS.Cells[GS.Col, GS.Row];
  GS.ElipsisButton := (GS.Col = 6);
end;

procedure TOM_PARREGLEFID.GSElipsisClick(Sender: TObject);
begin
  if GS.Col = 6 then GS.Cells[6,GS.Row] := GetArticleRecherche_Disp(Nil,'','FFO','',True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : GSCellExit
Mots clefs ... :
*****************************************************************}

procedure TOM_PARREGLEFID.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if (GS.Cells[ACol, ARow] <> StCellCur) and (TFSaisieList(Ecran).FTypeAction <> taConsult) then
  begin
    SeuilModifie := True;
    if TF.State = dsBrowse then TF.State := dsEdit;
  end;
end;

procedure TOM_PARREGLEFID.InsertLigneGS;
begin
  if (GS.Row = GS.Rowcount - 1) then
  begin
    GS.InsertRow(GS.Rowcount);
    GS.Col := 0;
  end;
end;

procedure TOM_PARREGLEFID.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Vide, Cancel: Boolean;
  ARow, ACol: Longint;
begin
  Vide := (Shift = []);
  ARow := GS.Row;
  ACol := GS.Col;    
  Cancel := False;
  case Key of
    VK_RETURN:
      if Vide then SendMessage(GS.Handle, WM_KeyDown, VK_DOWN, 0);
    VK_DOWN:
      begin
        if Vide and (ARow = GS.RowCount - 1) then
        begin
          GSCellExit(GS, ACol, ARow, Cancel);
          if not Cancel then InsertLigneGS;
        end;
      end;
    VK_TAB:
      begin
        GSCellExit(GS, ACol, ARow, Cancel);
        if Vide and (GS.Col = 7) and (ARow = GS.RowCount - 1) then
        begin
          if not Cancel then InsertLigneGS;
          GS.Row := GS.RowCount - 1;
        end;
      end;
    VK_DELETE: if (Shift = [ssCtrl]) then
      begin
        Key := 0;
        SupLigneClick(nil);
      end;
  end;
end;

procedure TOM_PARFIDELITE.UtilBonAchatOnClick;
var Etat: Boolean;
begin
  Exit;

  Etat := GetField('GFO_UTILBONACHAT') = 'X';

  SetControlVisible('GFO_MODEPAIE', not Etat);
  SetControlVisible('TGFO_MODEPAIE', not Etat);
  SetControlVisible('GFO_TYPEREMISE', Etat);
  SetControlVisible('TGFO_TYPEREMISE', Etat);
end;

////////////////////////////////////////////////////////////////////////

procedure AGLUtilBonAchatOnClick(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_PARFIDELITE) then TOM_PARFIDELITE(OM).UtilBonAchatOnClick
  else exit;
end;

initialization
  registerclasses([TOM_PARFIDELITE]);
  registerclasses([TOM_PARREGLEFID]);
  RegisterAglProc('UtilBonAchatOnClick', TRUE, 2, AGLUtilBonAchatOnClick);
end.
