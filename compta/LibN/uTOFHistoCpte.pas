{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/09/2002
Modifié le ... : 17/01/20007
Description .. : Source TOF de la FICHE : HISTOCPTE ()
Mots clefs ... : TOF;HISTOCPTE
*****************************************************************}
Unit uTOFHistoCpte;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     HPanel,
     HRichOle,    // THRichEditOle
     HRichEdt,    // THRichEditEdt
     Graphics,    // TCanvas
     Grids,       // TGridDrawState
{$IFDEF EAGLCLIENT}
     MainEagl,
{$ELSE}
     FE_Main,     // AGLLanceFiche
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Ent1,
     uTob,         // Tob
     uLibExercice, // CGetExerciceNMoins2
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ENDIF MODENT1}
     uLibWindows;


procedure CC_LanceFicheHistoCpte( vStQui : string ; vStCtx : string = 'GEN' );
function ChargeMemoCREVBLOCNOTE   ( vStNature,  vStCode,  vStExercice : string ) : string;
function SauveMemoCREVBLOCNOTE    ( vStNature,  vStCode,  vStExercice, vStData : string ) : Boolean;

Type

  TCtxAffHisto = ( CtxGeneral, CtxAPG );

  TOF_HISTOCPTE = Class (TOF)

    FListe            : THGrid;

    FPages            : TPageControl;

    FTabMillesime     : TTabsheet;
    FTabCycle         : TTabsheet;
    FTabSynthese      : TTabsheet;
    FTabAPG           : TTabsheet;
    FTabSyntheses     : TTabSheet;

    FMemoMillesime    : THRichEditOle;
    FMemoCycle        : THRichEditOle;
    FMemoSynthese     : THRichEditOle;
    FMemoAPG          : THRichEditOle;
    FMemoSyntheses    : THRichEditOle;

    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    FStCtx           : string;
    FStGeneral       : string;
    FStCycleRevision : string;

    FExoDateN2 : TExoDate;

    procedure OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

    procedure InitContexte;
    procedure InitFListe;

    procedure ChargeMemoMillesime ( vRow : integer );
    procedure ChargeMemoCycle     ( vRow : integer );
    procedure ChargeMemoSynthese  ( vRow : integer );
    procedure ChargeMemoAPG       ( vRow : integer );
    procedure ChargeMemoSyntheses ( vRow : integer );


  end ;

Implementation

uses
  uFImgListe,    // FImgListe
  HSysMenu;      // THSystemMenu

Const cColLibelle   = 1;
      cColDebit     = 2;
      cColCredit    = 3;
      cColSolde     = 4;

      cRowTitre     = 0;
      cRowN2        = 1;
      cRowPrecedent = 2;
      cRowEnCours   = 3;
      cRowSuivant   = 4;

      cImgBlocNote = 77;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : __/__/____
Modifié le ... : 21/03/2007
Description .. :
Suite ........ : Valeur Possible de VstCtx
Suite ........ : GEN = millesime + cycle + synthese cycle
Suite ........ : APG = appréciation générale du dossier
Suite ........ : surement par la suite (EXP, CCY, SCY etc...)
Mots clefs ... :
*****************************************************************}
procedure CC_LanceFicheHistoCpte( vStQui : string ; vStCtx : string = 'GEN' );
begin
  AGLLanceFiche('CP', 'CPHISTOCPTE', '', '', vStCtx + ';' + vStQui );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /
Description .. : Liste des valeurs possibles pour CBN_NATURE :
Suite ........ : GEN = Commentaire Millésimé
Suite ........ : CYC = Commentaire sur le cycle
Suite ........ : CYS = Synthèse du cycle
Suite ........ : SYN = Synthèse des cycles
Suite ........ : APG = Appréciation générale du dossier
Suite ........ : EXP =
Mots clefs ... :
*****************************************************************}
function ChargeMemoCREVBLOCNOTE( vStNature, vStCode, vStExercice : string ) : string;
var lQuery : TQuery;
begin
  Result := '';
  try
    lQuery := OpenSQl('SELECT CBN_BLOCNOTE FROM CREVBLOCNOTE WHERE ' +
                      'CBN_NATURE = "'+ vStNature + '" AND ' +
                      'CBN_CODE = "' + vStCode + '" AND ' +
                      'CBN_EXERCICE = "' + vStExercice + '"', True);

    if not lQuery.Eof then
      Result := lQuery.FindField('CBN_BLOCNOTE').AsString;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... : 26/03/2007
Description .. :
Suite ........ : Type possible pour CBN_NATURE :
Suite ........ : GEN = compte général
Suite ........ : CCY = commentaire du cycle
Suite ........ : SCY = synthèse du cycle
Suite ........ : SYN = synthèse des cycles
Suite ........ : APR = Appréciation Générale du dossier
Suite ........ : DRF = Détermination résultat fiscale
Mots clefs ... :
*****************************************************************}
function SauveMemoCREVBLOCNOTE( vStNature, vStCode, vStExercice, vStData : string ) : Boolean;
var lTob     : Tob;
    lTobTemp : Tob;
begin
  Result := True;
  try
    try
      lTob := Tob.Create('', nil, -1);
      lTob.LoadDetailDBFromSQL('CREVBLOCNOTE',
                               'SELECT * FROM CREVBLOCNOTE WHERE ' +
                               'CBN_NATURE = "' + vStNature + '" AND ' +
                               'CBN_CODE = "' + vStCode + '" AND ' +
                               'CBN_EXERCICE = "' + vStExercice + '"');

      if lTob.Detail.Count > 0 then
      begin
        lTobTemp := lTob.Detail[0];
        lTobTemp.PutValue('CBN_BLOCNOTE', vStData );
        lTobTemp.PutValue('CBN_DATEMODIF', Now );
        lTobTemp.PutValue('CBN_UTILISATEUR', V_Pgi.User );
        lTob.UpdateDB;
      end
      else
      begin
        lTobTemp := Tob.Create('CREVBLOCNOTE', lTob, -1);
        lTobTemp.PutValue('CBN_NATURE', vStNature );
        lTobTemp.PutValue('CBN_CODE', vStCode );
        lTobTemp.PutValue('CBN_EXERCICE', vStExercice );
        lTobTemp.PutValue('CBN_BLOCNOTE', vStData );
        lTobTemp.PutValue('CBN_DATEMODIF', Now );
        lTobTemp.PutValue('CBN_UTILISATEUR', V_Pgi.User );
        lTobTemp.PutValue('CBN_DATECREATION', Date );
        lTobTemp.PutValue('CBN_CREATEUR', V_Pgi.User );
        lTob.InsertDb(nil);
      end;

    except
      on E: Exception do
      begin
        Result := False;
        PgiError('Erreur de requête SQL : ' + E.Message, 'SauveCREVBLOCNOTE');
      end;
    end;

  finally
    FreeAndNil( lTob );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... : 26/03/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOCPTE.OnArgument (S : String ) ;
begin
  Inherited ;
  Ecran.HelpContext := 7601900;

  FListe               := THGrid(GetControl('FLISTE', True));
  FListe.OnRowEnter    := OnRowEnterFListe;
  FListe.GetCellCanvas := GetCellCanvasFListe;

  FPages               := TPageControl(GetControl('FPAGES', True));

  FTabMillesime        := TTabsheet(GetControl('FTABMILLESIME', True));
  FTabCycle            := TTabsheet(GetControl('FTABCYCLE', True));
  FTabSynthese         := TTabsheet(GetControl('FTABSYNTHESE', True));
  FTabAPG              := TTabsheet(GetControl('FTABAPG', True));
  FTabSyntheses        := TTabsheet(GetControl('FTABSYNTHESES', True));

  FMemoMillesime       := THRichEditOle(GetControl('FMEMOMILLESIME', True));
  FMemoCycle           := THRichEditOle(GetControl('FMEMOCYCLE', True));
  FMemoSynthese        := THRichEditOle(GetControl('FMEMOSYNTHESE', True));
  FMemoAPG             := THRichEditOle(GetControl('FMEMOAPG', True));
  FMemoSyntheses       := THRichEditOle(GetControl('FMEMOSYNTHESES', True));

  // Chargement de l'historique du compte
  FStCtx     := ReadTokenSt(S);
  FStGeneral := '';
  FStCycleRevision := '';

  if FStCtx = 'GEN' then
    FStGeneral := ReadTokenSt(S);

  if FStCtx = 'CCY' then
    FStCycleRevision := ReadTokenSt(S);

  FExoDateN2 := CGetExerciceNMoins2;

  // Placement systématique sur l'encours
  FListe.Row := cRowEnCours;
  if CtxPCl in V_Pgi.PgiContexte then
  begin
    if VH^.CPExoRef.Code = VH^.EnCours.Code then
      FListe.Row := cRowEnCours
    else
      if VH^.CpExoRef.Code = VH^.Suivant.Code then
        FListe.Row := cRowSuivant;
  end;

  InitContexte;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOCPTE.InitContexte;
begin
  InitFListe; // FSTCycleRevision dans INITFListe

  // GCO - 11/06/2007
  if (not V_Pgi.ModeTSE) then
  begin
    if FImgListe = nil then
      Application.CreateForm(TFImgListe, FImgListe);
    FPages.Images := FImgListe.BiblioActes;
  end;

  FTabMillesime.TabVisible := (FStCtx = 'GEN');
  FTabCycle.TabVisible     := (FStCycleRevision <> '') and (VH^.OkModRIC);
  FTabSynthese.TabVisible  := (FStCycleRevision <> '') and (VH^.OkModRIC);
  FTabAPG.TabVisible       := (FStCtx = 'APG') and (VH^.OkModRIC);
  FTabSyntheses.TabVisible := (FStCtx = 'APG') and (VH^.OkModRIC);

  if FTabMillesime.TabVisible then
  begin
    ChargeMemoMillesime( FListe.Row );
  end;

  if FTabCycle.TabVisible then
  begin
    FTabCycle.Caption := FTabCycle.Caption + ' ' + FStCycleRevision;
    ChargeMemoCycle   ( FListe.Row );
  end;

  if FTabSynthese.TabVisible then
  begin
    FTabSynthese.Caption := FTabSynthese.Caption + ' ' + FStCycleRevision;
    ChargeMemoSynthese ( FListe.Row );
  end;

  if FTabAPG.TabVisible then
  begin
    ChargeMemoAPG( FListe.Row );
  end;

  if FTabSyntheses.TabVisible then
  begin
    ChargeMemoSyntheses ( FListe.Row );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOCPTE.InitFListe;
var lQuery : tQuery;
begin
  try
    FListe.ColAligns[cColLibelle]  := TaLeftJustify;
    FListe.ColAligns[cColDebit]    := TaRightJustify;
    FListe.ColAligns[cColCredit]   := TaRightJustify;
    FListe.ColAligns[cColSolde]    := TaRightJustify;

    FListe.Cells[cColLibelle, cRowTitre] := TraduireMemoire('Exercice');
    FListe.Cells[cColDebit, cRowTitre]   := TraduireMemoire('Débit');
    FListe.Cells[cColCredit, cRowTitre]  := TraduireMemoire('Crédit');
    FListe.Cells[cColSolde, cRowTitre]   := TraduireMemoire('Solde');

    // Remplissage du FixedRow
    FListe.Cells[cColLibelle,cRowN2]        := TraduireMemoire('Exercice N-2') + IIF(FExoDateN2.Code <> '', ' du ' + DateToStr(FExoDateN2.Deb) + ' au ' + DatetoStr(FExoDateN2.Fin),'');
    FListe.Cells[cColLibelle,cRowPrecedent] := TraduireMemoire('Exercice précédent') + IIF(VH^.Precedent.Code <> '', ' du ' + DateToStr(VH^.Precedent.Deb) + ' au ' + DatetoStr(VH^.Precedent.Fin), '');
    FListe.Cells[cColLibelle,cRowEnCours]   := TraduireMemoire('Exercice en cours') + IIF(VH^.EnCours.Code <> '', ' du ' + DateToStr(VH^.EnCours.Deb) + ' au ' + DatetoStr(VH^.EnCours.Fin), '');
    FListe.Cells[cColLibelle,cRowSuivant]   := TraduireMemoire('Exercice suivant') + IIF(VH^.Suivant.Code <> '', ' du ' + DateToStr(VH^.Suivant.Deb) + ' au ' + DatetoStr(VH^.Suivant.Fin), '');

    if (FStCtx = 'APG') or (FStCtx = 'CCY') then
    begin
      if FStCtx = 'APG' then
        Ecran.Caption := TraduireMemoire('Historique de l''appréciation générale')
      else
        Ecran.Caption := TraduireMemoire('Historique des commentaires du cycle') + ' : ' +
                         FStCycleRevision + ' - '+
                         GetColonneSQL('CREVCYCLE', 'CCY_LIBELLECYCLE', 'CCY_CODECYCLE = "' + FStCycleRevision + '"');
      UpdateCaption( Ecran );

      FListe.ColWidths [cColDebit] := -1;
      FListe.ColLengths[cColDebit] := -1;

      FListe.ColWidths [cColCredit] := -1;
      FListe.ColLengths[cColCredit] := -1;

      FListe.ColWidths [cColSolde] := -1;
      FListe.ColLengths[cColSolde] := -1;
    end;

    if (FStCtx = 'GEN') and (FStGeneral <> '') then
    begin
      lQuery := OpenSql('SELECT G_LIBELLE, G_TOTDEBP, G_TOTCREP, G_TOTDEBE, ' +
                        'G_TOTCREE, G_TOTDEBS, G_TOTCRES, G_TOTDEBN2, ' +
                        'G_TOTCREN2, G_CYCLEREVISION FROM GENERAUX WHERE G_GENERAL = "' + FStGeneral + '"', True );

      if not lQuery.Eof then
      begin
        Ecran.Caption := TraduireMemoire('Historique du compte') + ' : ' + FStGeneral + ' ' + lQuery.FindField('G_Libelle').AsString;
        UpdateCaption( Ecran );
        FStCycleRevision := lQuery.FindField('G_CYCLEREVISION').AsString;

        FListe.Cells[cColDebit,cRowN2]   := StrFMontant(lQuery.FindField('G_TOTDEBN2').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColCredit,cRowN2]  := StrFMontant(lQuery.FindField('G_TOTCREN2').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColSolde,cRowN2]   := AfficheDBCR(lQuery.FindField('G_TOTDEBN2').AsFloat - lQuery.FindField('G_TOTCREN2').AsFloat);

        // Exercice précédent
        FListe.Cells[cColDebit,cRowPrecedent]   := StrfMontant(lQuery.FindField('G_TOTDEBP').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColCredit,cRowPrecedent]  := StrfMontant(lQuery.FindField('G_TOTCREP').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColSolde,cRowPrecedent]   := AfficheDBCR(lQuery.FindField('G_TOTDEBP').AsFloat - lQuery.FindField('G_TOTCREP').AsFloat);

        // Exercice En Cours
        FListe.Cells[cColDebit,cRowEnCours]   := StrfMontant(lQuery.FindField('G_TOTDEBE').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColCredit,cRowEnCours]  := StrfMontant(lQuery.FindField('G_TOTCREE').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColSolde,cRowEnCours]   := AfficheDBCR(lQuery.FindField('G_TOTDEBE').AsFloat - lQuery.FindField('G_TOTCREE').AsFloat);

        // Exercice Suivant
        FListe.Cells[cColDebit,cRowSuivant]   := StrfMontant(lQuery.FindField('G_TOTDEBS').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColCredit,cRowSuivant]  := StrfMontant(lQuery.FindField('G_TOTCRES').AsFloat, 15,V_PGI.OkDecV,'',True);
        FListe.Cells[cColSolde,cRowSuivant]   := AfficheDBCR(lQuery.FindField('G_TOTDEBS').AsFloat - lQuery.FindField('G_TOTCRES').AsFloat);
      end;
    end;
  finally
    Ferme( lQuery );
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOCPTE.OnLoad ;
begin
  Inherited ;
end ;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/04/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOCPTE.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if FTabMillesime.TabVisible then ChargeMemoMillesime( Ou );
  if FTabCycle.TabVisible     then ChargeMemoCycle( Ou );
  if FTabSynthese.TabVisible  then ChargeMemosynthese ( Ou );
  if FTabAPG.TabVisible       then ChargeMemoAPG( Ou );
  if FTabSyntheses.TabVisible then ChargeMemoSyntheses( Ou );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOCPTE.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then Exit;
  if (FListe.Row <> ARow) then
  begin
    case ARow of
      1 : Canvas.Font.Color := clNavy;
      2 : Canvas.Font.Color := clGreen;
      3 : Canvas.Font.Color := ClBlack;
      4 : Canvas.Font.Color := clMaroon;
    else
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/04/2005
Modifié le ... : 17/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOCPTE.ChargeMemoMillesime(vRow: integer);
var lStTemp : string;
begin
  FMemoMillesime.Clear;
  lStTemp := '';
  case vRow of
    1 : lStTemp := ChargeMemoCREVBLOCNOTE('GEN', FStGeneral, FExoDateN2.Code);
    2 : lStTemp := ChargeMemoCREVBLOCNOTE('GEN', FStGeneral, VH^.Precedent.Code);
    3 : lStTemp := ChargeMemoCREVBLOCNOTE('GEN', FStGeneral, VH^.EnCours.Code);
    4 : lStTemp := ChargeMemoCREVBLOCNOTE('GEN', FStGeneral, VH^.Suivant.Code);
  else
  end;

  StringToRich(FMemoMillesime, lStTemp );

  FTabMillesime.ImageIndex := -1;
  if FMemoMillesime.Text <> '' then
  begin
    if V_Pgi.ModeTse then
      FTabMillesime.Caption := FTabMillesime.Caption +  ' (***)'
    else
      FTabMillesime.ImageIndex := cImgBlocNote;
  end;
  
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOCPTE.ChargeMemoCycle(vRow: integer);
var lStTemp : string;
begin
  FMemoCycle.Clear;
  if FStCycleRevision = '' then Exit;
  lStTemp := '';
  case vRow of
    1 : lStTemp := ChargeMemoCREVBLOCNOTE('CCY', FStCycleRevision, FExoDateN2.Code);
    2 : lStTemp := ChargeMemoCREVBLOCNOTE('CCY', FStCycleRevision, VH^.Precedent.Code);
    3 : lStTemp := ChargeMemoCREVBLOCNOTE('CCY', FStCycleRevision, VH^.EnCours.Code);
    4 : lStTemp := ChargeMemoCREVBLOCNOTE('CCY', FStCycleRevision, VH^.Suivant.Code);
  else
  end;
  StringToRich(FMemoCycle, lStTemp);

  FTabCycle.ImageIndex := -1;
  if FMemoCycle.Text <> '' then
  begin
    if V_Pgi.ModeTse then
      FTabCycle.Caption := FTabCycle.Caption +  ' (***)'
    else
      FTabCycle.ImageIndex := cImgBlocNote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_HISTOCPTE.ChargeMemoSynthese(vRow: integer);
var lStTemp : string;
begin
  FMemoSynthese.Clear;
  if FStCycleRevision = '' then Exit;
  lStTemp := '';
  case vRow of
    1 : lStTemp := ChargeMemoCREVBLOCNOTE('SCY', FStCycleRevision, FExoDateN2.Code);
    2 : lStTemp := ChargeMemoCREVBLOCNOTE('SCY', FStCycleRevision, VH^.Precedent.Code);
    3 : lStTemp := ChargeMemoCREVBLOCNOTE('SCY', FStCycleRevision, VH^.EnCours.Code);
    4 : lStTemp := ChargeMemoCREVBLOCNOTE('SCY', FStCycleRevision, VH^.Suivant.Code);
  else
  end;
  StringToRich(FMemoSynthese, lStTemp);

  FTabSynthese.ImageIndex := -1;
  if FMemoSynthese.Text <> '' then
  begin
    if V_Pgi.ModeTse then
      FTabSynthese.Caption := FTabSynthese.Caption +  ' (***)'
    else
      FTabSynthese.ImageIndex := cImgBlocNote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOCPTE.ChargeMemoAPG( vRow : integer );
begin
  FMemoAPG.Clear;
  case vRow of
    1 : StringToRich(FMemoAPG, ChargeMemoCREVBLOCNOTE('APG', '', FExoDateN2.Code));
    2 : StringToRich(FMemoAPG, ChargeMemoCREVBLOCNOTE('APG', '', VH^.Precedent.Code));
    3 : StringToRich(FMemoAPG, ChargeMemoCREVBLOCNOTE('APG', '', VH^.EnCours.Code));
    4 : StringToRich(FMemoAPG, ChargeMemoCREVBLOCNOTE('APG', '', VH^.Suivant.Code));
  else
  end;

  FTabAPG.ImageIndex := -1;
  if FMemoAPG.Text <> '' then
  begin
    if V_Pgi.ModeTse then
      FTabAPG.Caption := FTabAPG.Caption +  ' (***)'
    else
      FTabAPG.ImageIndex := cImgBlocNote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOCPTE.ChargeMemoSyntheses( vRow : integer );
begin
  FMemoSyntheses.Clear;
  case vRow of
    1 : StringToRich(FMemoSyntheses, ChargeMemoCREVBLOCNOTE('SYN', '', FExoDateN2.Code));
    2 : StringToRich(FMemoSyntheses, ChargeMemoCREVBLOCNOTE('SYN', '', VH^.Precedent.Code));
    3 : StringToRich(FMemoSyntheses, ChargeMemoCREVBLOCNOTE('SYN', '', VH^.EnCours.Code));
    4 : StringToRich(FMemoSyntheses, ChargeMemoCREVBLOCNOTE('SYN', '', VH^.Suivant.Code));
  else
  end;

  FTabSyntheses.ImageIndex := -1;
  if FMemoSyntheses.Text <> '' then
  begin
    if V_Pgi.ModeTse then
      FTabSyntheses.Caption := FTabSyntheses.Caption +  ' (***)'
    else
      FTabSyntheses.ImageIndex := cImgBlocNote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/02/2005
Modifié le ... : 11/06/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_HISTOCPTE.OnClose ;
begin
  FreeAndNil( FImgListe );
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////////


Initialization
  registerclasses ( [ TOF_HISTOCPTE ] ) ;
end.

