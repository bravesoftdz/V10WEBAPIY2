{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Source TOF de la FICHE : MFOSOLDECLIENT ()
Suite ........ : Affichage du solde du client
Suite ........ :
Suite ........ : Source TOF de la FICHE : MFODETAILSOLDECLI ()
Suite ........ : Affichage du détail du solde du client
Mots clefs ... : TOF;MFOSOLDECLIENT;MFODETAILSOLDECLI
*****************************************************************}
unit MFOSOLDECLIENT_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  db, dbtables, Fe_Main,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
  Graphics, UTOB, Vierge, HTB97, AGLInit;

type
  TOF_MFOSOLDECLIENT = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    CodeTiers: string; // code du tiers
    MtSoldeEtab: double; // Solde dans l'établissement
    MtSoldeAutre: double; // Solde dans les autres établissements
    TOBEtab: TOB; // TOB des cumuls de l'établissement
    TOBAutre: TOB; // TOB des cumuls des autres établissements
    BSOLDEETAB: TToolbarButton97; // bouton pour voir le solde de tous les établissements
    PETABHaut: integer; // hauteur du panel PETAB
    GSETABHaut: integer; // hauteur de la grille GSETAB
    procedure AlloueTOB;
    procedure LibereTOB;
    procedure CalculSolde(Autres: boolean);
    procedure BZoomClick(Sender: TObject);
    procedure BSOLDEETABClick(Sender: TObject);
  end;

type
  TOF_MFODETAILSOLDECLI = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    GS: THGrid; // grille d'afficahge
    procedure BZoomClick(Sender: TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
  end;

function FOCalculSoldeClient(CodeTiers: string; TousEtab: boolean): double;

implementation

uses
  Ent1, EntGC, FOUtil, TickUtilFO, MC_Lib, UtilPgi;

{==============================================================================================}
{============================ CALCUL DU SOLDE DU CLIENT =======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Chargement des opérations de caisse du client
Mots clefs ... :
*****************************************************************}

procedure ChargeOperCaisse(CodeTiers: string; TOBOpc: TOB);
var
  sSql: string;
  QQ: TQuery;
begin
  if TOBOpc = nil then Exit;

  sSql := 'SELECT GL_ETABLISSEMENT,GL_ARTICLE,GL_TOTALTTC,'
    + ' GL_DATEPIECE,GL_NUMERO,GA_LIBELLE,GA_TYPEARTFINAN,'
    + ' GOC_NUMPIECELIEN,GOC_TOTALDISPOTTC,GOC_NUMBON'
    + ' FROM LIGNE LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE'
    + ' LEFT JOIN OPERCAISSE ON GOC_NATUREPIECEG=GL_NATUREPIECEG AND GOC_SOUCHE=GL_SOUCHE'
    + ' AND GOC_NUMERO=GL_NUMERO AND GOC_INDICEG=GL_INDICEG'
    + ' AND GOC_PREFIXE="GL" AND GOC_NUMORDRE=GL_NUMORDRE'
    + ' WHERE GL_TIERS="' + CodeTiers + '" AND GL_TYPEARTICLE="FI"';
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then TOBOpc.LoadDetailDB('LIGNE', '', '', QQ, False);
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Cumul des opérations de caisse du client
Mots clefs ... :
*****************************************************************}

procedure CumulOperCaisse(CodeTiers: string; TOBOpc, TOBEtab, TOBAutre: TOB);
var
  TypeArt, Stg, NumLien: string;
  MtVendu, MtRegle, MtVd, MtRg: double;
  TOBL, TOBC, TOBM: TOB;
  Ind: integer;
begin
  if (TOBOpc = nil) or (TOBEtab = nil) or (TOBAutre = nil) then Exit;

  for Ind := TOBOpc.Detail.Count - 1 downto 0 do
  begin
    TOBL := TOBOpc.Detail[Ind];
    if TOBL.GetValue('GL_ETABLISSEMENT') = VH^.EtablisDefaut then
      TOBM := TOBEtab
    else
      TOBM := TOBAutre;
    TypeArt := TOBL.GetValue('GA_TYPEARTFINAN');
    TOBC := TOBM.FindFirst(['CODE'], [TypeArt], False);
    if TOBC = nil then
    begin
      TOBC := TOB.Create('', TOBM, -1);
      TOBC.AddChampSupValeur('TYPE', 'GL');
      TOBC.AddChampSupValeur('CODE', TypeArt);
      Stg := RechDom('GCTYPEARTFINAN', TypeArt, False);
      TOBC.AddChampSupValeur('LIBELLE', Stg);
      MtVendu := 0;
      TOBC.AddChampSupValeur('MTVENDU', MtVendu);
      MtRegle := 0;
      TOBC.AddChampSupValeur('MTREGLE', MtRegle);
    end else
    begin
      MtRegle := TOBC.GetValue('MTREGLE');
    end;
    NumLien := vString(TOBL.GetValue('GOC_NUMPIECELIEN'));
    if NumLien = '' then TOBL.PutValue('GOC_NUMPIECELIEN', NumLien);
    if VarIsNull(TOBL.GetValue('GOC_NUMBON')) then TOBL.PutValue('GOC_NUMBON', '');
    MtVd := TOBL.GetValue('GL_TOTALTTC');
    MtRg := vDouble(TOBL.GetValue('GOC_TOTALDISPOTTC'));
    if MtRg = 0 then TOBL.PutValue('GOC_TOTALDISPOTTC', MtRg);

    if FOIsOpCaisseCumul('GL', TypeArt, MtVd) then
    begin
      // cas d'un cumul sans liaison des règlements
      if (NumLien = '') and (MtRg <> MtVd) then
        MtRg := MtVd;
    end else
      if FOIsOpCaisseLiee('GL', TypeArt, MtVd) then
    begin
      // cas d'un règlement qui doit être lié à un cumul mais saisi sans liaison
      if (NumLien = '') and (MtRg = 0) then
        MtRg := MtVd;
    end;

    MtRegle := Arrondi(MtRegle + MtRg, V_PGI.OkDecV);
    TOBC.PutValue('MTREGLE', MtRegle);
    TOBL.ChangeParent(TOBC, -1);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Chargement des échéances du client
Mots clefs ... :
*****************************************************************}

procedure ChargeEcheances(CodeTiers: string; TOBMdr: TOB);
var
  sSql: string;
  QQ: TQuery;
begin
  if TOBMdr = nil then Exit;

  sSql := 'SELECT GP_ETABLISSEMENT,GPE_MODEPAIE,GPE_MONTANTECHE,GPE_DATEPIECE,'
    + ' GPE_NUMERO,MP_LIBELLE,MP_TYPEMODEPAIE,'
    + ' GOC_NUMPIECELIEN,GOC_TOTALDISPOTTC,GOC_NUMBON'
    + ' FROM PIEDECHE LEFT JOIN MODEPAIE ON MP_MODEPAIE=GPE_MODEPAIE'
    + ' LEFT JOIN PIECE ON GP_NATUREPIECEG=GPE_NATUREPIECEG AND GP_SOUCHE=GPE_SOUCHE'
    + ' AND GP_NUMERO=GPE_NUMERO AND GP_INDICEG=GPE_INDICEG'
    + ' LEFT JOIN OPERCAISSE ON GOC_NATUREPIECEG=GPE_NATUREPIECEG AND GOC_SOUCHE=GPE_SOUCHE'
    + ' AND GOC_NUMERO=GPE_NUMERO AND GOC_INDICEG=GPE_INDICEG'
    + ' AND GOC_PREFIXE="GPE" AND GOC_NUMLIGNE=GPE_NUMECHE'
    + ' WHERE GPE_TIERS="' + CodeTiers + '"';
  QQ := OpenSQL(sSql, True);
  if not QQ.EOF then TOBMdr.LoadDetailDB('PIEDECHE', '', '', QQ, False);
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Chargement des échéances du client
Mots clefs ... :
*****************************************************************}

procedure CumulEcheances(CodeTiers: string; TOBMdr, TOBEtab, TOBAutre: TOB);
var
  TypeMdp, Stg, NumLien: string;
  MtVendu, MtRegle, MtVd, MtRg: double;
  TOBL, TOBC, TOBM: TOB;
  Ind: integer;
begin
  if (TOBMdr = nil) or (TOBEtab = nil) or (TOBAutre = nil) then Exit;
  for Ind := TOBMdr.Detail.Count - 1 downto 0 do
  begin
    TOBL := TOBMdr.Detail[Ind];
    if TOBL.GetString('GP_ETABLISSEMENT') = VH^.EtablisDefaut then
      TOBM := TOBEtab
    else
      TOBM := TOBAutre;
    TypeMdp := TOBL.GetString('MP_TYPEMODEPAIE');
    TOBC := TOBM.FindFirst(['CODE'], [TypeMdp], False);
    if TOBC = nil then
    begin
      TOBC := TOB.Create('', TOBM, -1);
      TOBC.AddChampSupValeur('TYPE', 'GPE');
      TOBC.AddChampSupValeur('CODE', TypeMdp);
      Stg := RechDom('GCTYPEMODEPAIE', TypeMdp, False);
      TOBC.AddChampSupValeur('LIBELLE', Stg);
      MtVendu := 0;
      TOBC.AddChampSupValeur('MTVENDU', MtVendu);
      MtRegle := 0;
      TOBC.AddChampSupValeur('MTREGLE', MtRegle);
    end else
    begin
      MtVendu := TOBC.GetDouble('MTVENDU');
      MtRegle := TOBC.GetDouble('MTREGLE');
    end;
    NumLien := vString(TOBL.GetString('GOC_NUMPIECELIEN'));
    if NumLien = '' then
      TOBL.SetString('GOC_NUMPIECELIEN', NumLien);
    if VarIsNull(TOBL.GetString('GOC_NUMBON')) then
      TOBL.SetString('GOC_NUMBON', '');
    MtVd := TOBL.GetDouble('GPE_MONTANTECHE');
    MtVendu := Arrondi(MtVendu + MtVd, V_PGI.OkDecV);
    TOBC.SetDouble('MTVENDU', MtVendu);
    MtRg := vDouble(TOBL.GetDouble('GOC_TOTALDISPOTTC'));
    if MtRg = 0 then TOBL.SetDouble('GOC_TOTALDISPOTTC', MtRg);
    if FOIsOpCaisseCumul('GPE', TypeMdp, MtVd) then
    begin
      // cas d'un cumul sans liaison des règlements
      if (NumLien = '') and (MtRg <> MtVd) then
        MtRg := MtVd;
    end else
      if FOIsOpCaisseLiee('GPE', TypeMdp, MtVd) then
    begin
      // cas d'un règlement qui doit être lié à un cumul mais saisi sans liaison
      if (NumLien = '') and (MtRg = 0) then
        MtRg := MtVd;
    end;

    MtRegle := Arrondi(MtRegle + MtVd - MtRg, V_PGI.OkDecV);
    TOBC.SetDouble('MTREGLE', MtRegle);
    TOBL.ChangeParent(TOBC, -1);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Cumul des échéances ou des opérations de caisse du client
Mots clefs ... :
*****************************************************************}

function CumulLignes(TOBLignes: TOB): double;
var
  TOBL: TOB;
  Ind: integer;
  TotVendu, TotRegle, dVal: double;
begin
  Result := 0;
  if TOBLignes = nil then Exit;
  TotVendu := 0;
  TotRegle := 0;
  for Ind := 0 to TOBLignes.Detail.Count - 1 do
  begin
    TOBL := TOBLignes.Detail[Ind];
    dVal := TOBL.GetValue('MTVENDU');
    TotVendu := Arrondi(TotVendu + dVal, V_PGI.OkDecV);
    dVal := TOBL.GetValue('MTREGLE');
    TotRegle := Arrondi(TotRegle + dVal, V_PGI.OkDecV);
  end;
  Result := Arrondi(TotRegle - TotVendu, V_PGI.OkDecV);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Calcul du solde d'un client
Mots clefs ... :
*****************************************************************}

function FOCalculSoldeClient(CodeTiers: string; TousEtab: boolean): double;
var
  TOBOpc, TOBMdr, TOBEtab, TOBAutre: TOB;
begin
  TOBEtab := TOB.Create('', nil, -1);
  TOBAutre := TOB.Create('', nil, -1);
  TOBOpc := TOB.Create('', nil, -1);
  TOBMdr := TOB.Create('', nil, -1);

  {$IFDEF AGL580}
  // V_PGI.SuppressionDesNullsDeLaTob := False ;
  {$ENDIF AGL580}
  ChargeOperCaisse(CodeTiers, TOBOpc);
  CumulOperCaisse(CodeTiers, TOBOpc, TOBEtab, TOBAutre);
  ChargeEcheances(CodeTiers, TOBMdr);
  CumulEcheances(CodeTiers, TOBMdr, TOBEtab, TOBAutre);
  {$IFDEF AGL580}
  // V_PGI.SuppressionDesNullsDeLaTob := True;
  {$ENDIF AGL580}
  Result := CumulLignes(TOBEtab);
  if TousEtab then Result := Result + CumulLignes(TOBAutre);

  FreeAndNil(TOBMdr);
  FreeAndNil(TOBOpc);
  FreeAndNil(TOBEtab);
  FreeAndNil(TOBAutre);
end;

{==============================================================================================}
{================================ TOF_MFOSOLDECLIENT ==========================================}
{==============================================================================================}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Création des TOB
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.AlloueTOB;
begin
  TOBEtab := TOB.Create('', nil, -1);
  TOBAutre := TOB.Create('', nil, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Libération des TOB
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.LibereTOB;
begin
  if TOBEtab <> nil then FreeAndNil(TOBEtab);
  if TOBAutre <> nil then FreeAndNil(TOBAutre);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Calcul du solde du client
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.CalculSolde(Autres: boolean);
var
  TOBM: TOB;
  Ctrl: TControl;
  dVal: double;
begin
  if Autres then
  begin
    Ctrl := GetControl('GSAUTRE');
    TOBM := TOBAutre;
  end else
  begin
    Ctrl := GetControl('GSETAB');
    TOBM := TOBEtab;
  end;
  if (Ctrl <> nil) and (Ctrl is THGrid) then
  begin
    TOBM.Detail.Sort('CODE');
    TOBM.PutGridDetail(THGrid(Ctrl), True, False, 'LIBELLE;MTVENDU;MTREGLE', True);
    TFVierge(Ecran).HMTrad.ResizeGridColumns(THGrid(Ctrl));
    THGrid(Ctrl).UpdateTitres;
    THGrid(Ctrl).ColAligns[1] := taRightJustify;
    THGrid(Ctrl).ColAligns[2] := taRightJustify;
  end;

  dVal := CumulLignes(TOBM);
  if Autres then
  begin
    Ctrl := GetControl('MTSOLDEAUTRE');
    MtSoldeAutre := dVal;
  end else
  begin
    Ctrl := GetControl('MTSOLDEETAB');
    MtSoldeEtab := dVal;
  end;
  if (Ctrl <> nil) and (Ctrl is THNumEdit) then
  begin
    THNumEdit(Ctrl).Value := dVal;
    if dVal < 0 then THNumEdit(Ctrl).Font.Color := clRed;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Affichage du détail de la ligne
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.BZoomClick(Sender: TObject);
var
  Ind: integer;
begin
  TheTOB := nil;
  if (Screen.ActiveControl.Name = 'GSETAB') and (Screen.ActiveControl is THGrid) then
  begin
    Ind := THGrid(Screen.ActiveControl).Row - THGrid(Screen.ActiveControl).FixedRows;
    if (Ind >= 0) and (Ind < TOBEtab.Detail.Count) then
      TheTOB := TOBEtab.Detail[Ind];
  end else
    if (Screen.ActiveControl.Name = 'GSAUTRE') and (Screen.ActiveControl is THGrid) then
  begin
    Ind := THGrid(Screen.ActiveControl).Row - THGrid(Screen.ActiveControl).FixedRows;
    if (Ind >= 0) and (Ind < TOBAutre.Detail.Count) then
      TheTOB := TOBAutre.Detail[Ind];
  end;

  if TheTOB <> nil then
  begin
    AglLanceFiche('MFO', 'MFODETAILSOLDECLI', '', '', '');
    TheTOB := nil;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Affichage du solde de tous les établissements
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.BSOLDEETABClick(Sender: TObject);
var
  Ctrl: TControl;
  dVal: double;
begin
  if BSOLDEETAB = nil then Exit;
  if BSOLDEETAB.Down then
  begin
    // solde de tous les établissements
    dVal := (MtSoldeAutre + MtSoldeEtab);
    SetControlProperty('PETAB', 'Align', alTop);
    SetControlProperty('PETAB', 'Height', PETABHaut);
    SetControlProperty('PAUTRE', 'Align', alClient);
    SetControlVisible('PAUTRE', True);
    SetControlVisible('MTSOLDEETAB', True);
    SetControlProperty('GSETAB', 'Align', alNone);
    SetControlProperty('GSETAB', 'Height', GSETABHaut);
  end else
  begin
    // solde de l'établissement courant
    dVal := MtSoldeEtab;
    SetControlVisible('PAUTRE', False);
    SetControlProperty('PAUTRE', 'Align', alNone);
    SetControlProperty('PETAB', 'Align', alClient);
    SetControlVisible('MTSOLDEETAB', False);
    SetControlProperty('GSETAB', 'Align', alClient);
  end;

  Ctrl := GetControl('MTSOLDE');
  if (Ctrl <> nil) and (Ctrl is THNumEdit) then
  begin
    THNumEdit(Ctrl).Value := dVal;
    if dVal < 0 then THNumEdit(Ctrl).Font.Color := clRed;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnNew;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnDelete;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnUpdate;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnLoad;
var
  TOBOpc, TOBMdr: TOB;
  Ctrl: TControl;
begin
  inherited;
  TOBOpc := TOB.Create('', nil, -1);
  TOBMdr := TOB.Create('', nil, -1);
  {$IFDEF AGL580}
  // V_PGI.SuppressionDesNullsDeLaTob := false;
  {$ENDIF AGL580}
  ChargeOperCaisse(CodeTiers, TOBOpc);
  CumulOperCaisse(CodeTiers, TOBOpc, TOBEtab, TOBAutre);
  ChargeEcheances(CodeTiers, TOBMdr);
  CumulEcheances(CodeTiers, TOBMdr, TOBEtab, TOBAutre);
  {$IFDEF AGL580}
  // V_PGI.SuppressionDesNullsDeLaTob := True;
  {$ENDIF AGL580}
  FreeAndNil(TOBMdr);
  FreeAndNil(TOBOpc);

  CalculSolde(False);
  CalculSolde(True);
  if (BSOLDEETAB <> nil) and ((TOBAutre.Detail.Count <= 0) or
    not (FOJaiLeDroit(85, False, False))) then
  begin
    BSOLDEETAB.Visible := False;
    BSOLDEETAB.Down := False;
  end;
  Ctrl := GetControl('PETAB');
  if Ctrl <> nil then PETABHaut := Ctrl.Height;
  Ctrl := GetControl('GSETAB');
  if Ctrl <> nil then GSETABHaut := Ctrl.Height;
  BSOLDEETABClick(nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnArgument(S: string);
var
  Stg, NomChamp, NomTiers: string;
  Ind: integer;
  Ctrl: TControl;
begin
  inherited;
  MtSoldeEtab := 0;
  MtSoldeAutre := 0;
  CodeTiers := '';
  NomTiers := '';
  Stg := S;
  repeat
    Stg := UpperCase(Trim(ReadTokenSt(S)));
    if Stg <> '' then
    begin
      Ind := Pos('=', Stg);
      if Ind <> 0 then
      begin
        NomChamp := Copy(Stg, 1, (Ind - 1));
        if NomChamp = 'T_TIERS' then
          CodeTiers := Trim(Copy(Stg, (Ind + 1), MaxInt))
        else if NomChamp = 'T_LIBELLE' then
          NomTiers := Trim(Copy(Stg, (Ind + 1), MaxInt));
      end;
    end;
  until S = '';
  if CodeTiers = '' then Close;

  if NomTiers = '' then NomTiers := RechDom('GCTIERS', CodeTiers, False);
  Stg := Ecran.Caption + ' ' + CodeTiers + ' ' + NomTiers;
  Ecran.Caption := Stg;

  Ctrl := GetControl('BSOLDEETAB');
  if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
  begin
    BSOLDEETAB := TToolbarButton97(Ctrl);
    BSOLDEETAB.OnClick := BSOLDEETABClick;
  end;
  Ctrl := GetControl('BZOOM');
  if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
    TToolbarButton97(Ctrl).OnClick := BZoomClick;
  Ctrl := GetControl('GSETAB');
  if (Ctrl <> nil) and (Ctrl is THGrid) then
    THGrid(Ctrl).OnDblClick := BZoomClick;
  Ctrl := GetControl('GSAUTRE');
  if (Ctrl <> nil) and (Ctrl is THGrid) then
    THGrid(Ctrl).OnDblClick := BZoomClick;

  AlloueTOB;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnClose;
begin
  inherited;
  LibereTOB;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnDisplay
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnDisplay();
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MFOSOLDECLIENT.OnCancel();
begin
  inherited;
end;

{==============================================================================================}
{=============================== TOF_MFODETAILSOLDECLI ========================================}
{==============================================================================================}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : Affichage du détail de la ligne
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.BZoomClick(Sender: TObject);
var
  Stg, sType, Titre, Msg: string;
  Ind: integer;
  CD: R_CleDoc;
begin
  if LaTOB = nil then Exit;

  if LaTOB.FieldExists('TYPE') then sType := LaTOB.GetValue('TYPE');
  Ind := THGrid(Screen.ActiveControl).Row - THGrid(Screen.ActiveControl).FixedRows;
  if (Ind >= 0) and (Ind < LaTOB.Detail.Count) then
  begin
    Stg := LaTOB.Detail[Ind].GetValue('GOC_NUMPIECELIEN');
    if Stg = '' then
    begin
      Titre := '';
      Msg := 'Pas de règlement lié';
    end else
    begin
      if ReadTokenSt(Stg) = 'GL' then
        Titre := 'Référence de l''opération de caisse liée'
      else
        Titre := 'Référence du règlement lié';
      DecodeRefPiece(Stg, CD);
      Msg := TraduireMemoire(RechDom('GCNATUREPIECEG', CD.NaturePiece, False))
        + TraduireMemoire(' n°') + IntToStr(CD.NumeroPiece)
        + TraduireMemoire(' du ') + DateToStr(CD.DatePiece)
        + TraduireMemoire(' ligne n°') + IntToStr(CD.NumLigne);
    end;
    PGIInfo(Msg, Titre);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnRowEnter sur la grille
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  Stg, sType: string;
  Ind: integer;
  Ok: boolean;
begin
  if LaTOB = nil then Exit;
  Ok := False;
  if LaTOB.FieldExists('TYPE') then sType := LaTOB.GetValue('TYPE');
  Ind := GS.Row - GS.FixedRows;
  if (Ind >= 0) and (Ind < LaTOB.Detail.Count) then
  begin
    Stg := LaTOB.Detail[Ind].GetValue('GOC_NUMPIECELIEN');
    Ok := (Stg <> '');
  end;
  SetControlEnabled('BZOOM', Ok);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnNew;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnDelete;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnUpdate;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnLoad;
var
  sType, sCode, sCol, sTitre, ColSort: string;
  Ctrl: TControl;
  Cancel: boolean;
begin
  inherited;
  if LaTOB = nil then Exit;
  sType := '';
  if LaTOB.FieldExists('TYPE') then sType := LaTOB.GetValue('TYPE');
  if sType = 'GL' then
  begin
    sCol := 'GL_ETABLISSEMENT;GL_DATEPIECE;GL_NUMERO;GA_LIBELLE;GOC_NUMBON;GL_TOTALTTC;GOC_TOTALDISPOTTC';
    ColSort := 'GL_NUMERO';
  end else
    if sType = 'GPE' then
  begin
    sCol := 'GP_ETABLISSEMENT;GPE_DATEPIECE;GPE_NUMERO;MP_LIBELLE;GOC_NUMBON;GPE_MONTANTECHE;GOC_TOTALDISPOTTC';
    ColSort := 'GPE_NUMERO';
  end else
    Exit;
  sTitre := Ecran.Caption + ' ' + LaTOB.GetValue('LIBELLE');
  Ecran.Caption := sTitre;
  sCode := LaTOB.GetValue('CODE');

  Ctrl := GetControl('GS');
  if (Ctrl <> nil) and (Ctrl is THGrid) then
  begin
    GS := THGrid(Ctrl);
    LaTOB.Detail.Sort(ColSort);
    LaTOB.PutGridDetail(GS, True, False, sCol, True);
    GS.UpdateTitres;
    GS.ColAligns[0] := taCenter;
    GS.ColAligns[1] := taCenter;
    GS.ColAligns[2] := taCenter;
    GS.ColAligns[5] := taRightJustify;
    GS.ColAligns[6] := taRightJustify;
    if not FOIsOpCaisseNoBon(sType, sCode, -1) then
    begin
      GS.ColWidths[4] := 0;
      GS.ColLengths[4] := -1;
    end;
    GS.OnDblClick := BZoomClick;
    GS.OnRowEnter := GSRowEnter;
    GSRowEnter(GS, 0, Cancel, True);
    TFVierge(Ecran).HMTrad.ResizeGridColumns(GS);
  end;

  Ctrl := GetControl('BZOOM');
  if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
    TToolbarButton97(Ctrl).OnClick := BZoomClick;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnArgument(S: string);
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnClose;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnDisplay
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnDisplay();
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 17/04/2003
Modifié le ... : 17/04/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MFODETAILSOLDECLI.OnCancel();
begin
  inherited;
end;

initialization
  registerclasses([TOF_MFOSOLDECLIENT]);
  registerclasses([TOF_MFODETAILSOLDECLI]);
end.
