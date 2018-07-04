{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CREVPARAMCYCLE (CREVPARAMCYCLE)
Mots clefs ... : TOM;CREVPARAMCYCLE
*****************************************************************}
Unit CREVPARAMCYCLE_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     eFiche,        // TFFiche
     MainEagl,      // AGLLanceFiche
{$ELSE}
     db,
     HDb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fiche,         // TFFiche
     FE_main,       // AGLLanceFiche
{$ENDIF}
     Forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     AGLInit,       // ActionToStr
     HQry,          // RecupWhereCritere
     HTb97,         // TToolBarButton97
     HSysMenu,      // HSystemMenu
     UTob ;

Type

  TOM_CREVPARAMCYCLE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;


  private
    FStPlanAssocie    : HString; // Liste des plans associés aux cycles
    FStListeCompte    : HString; // Liste des comptes
    FStListeExclusion : HString; // Liste des comptes exclus

    FBoLoading : Boolean; // Chargement du record en cours

    //fOnSaveElipsisClick : TNotifyEvent;

    procedure ChargeListePlan;
    procedure ChargeFListe;
    procedure ChargeActivation;
    procedure SauveFListe;
    procedure SauveActivation;

    function  TrouveDataType(vStValue : string) : string;
    procedure ChargeEnregistrement;

    function  MAJDateModifCREVPLAN : boolean;

  public
    Pages         : THPageControl2;
    FListe        : THGrid;
    ListePlan     : TListBox;

    BPlanRevision : TToolBarbutton97;
    BCopier       : TToolBarbutton97;
    BEfface       : TToolBarbutton97;
    BDelete       : TToolBarbutton97;

    Z_C1          : THValComboBox;
    ZO1           : THValComboBox;
    ZV1           : THEdit;
    ZG1           : THValComboBox;

    Z_C2          : THValComboBox;
    ZO2           : THValComboBox;
    ZV2           : THEdit;
    ZG2           : THValComboBox;

    Z_C3          : THValComboBox;
    ZO3           : THValComboBox;
    ZV3           : THEdit;

    procedure OnClickBPlanRevision ( Sender : TObject );
    procedure OnClickBCopier       ( Sender : TObject );
    procedure OnClickBEfface       ( Sender : TObject );

    procedure OnDblClickFListe     ( Sender : TObject );
    procedure OnCellExitFListe     ( Sender : TObject; var ACol,ARow: Longint; var Cancel: Boolean);
    procedure OnKeyPressFListe     ( Sender : TObject; var Key: Char);
    procedure OnKeyDownFListe      ( Sender : TObject; var Key: Word; Shift: TShiftState);

    procedure OnChangeZC           ( Sender : TObject );
    procedure OnChangeZxx          ( Sender : TObject );
    procedure OnElipsisClickZV     ( Sender : TObject );

    procedure OnKeyPressCPC_CODECYCLE(Sender: TObject; var Key: Char);

  end ;

////////////////////////////////////////////////////////////////////////////////
function CPLanceFiche_CREVPARAMCYCLE( vStLequel : string; vStActionFiche : TActionFiche ) : string;
////////////////////////////////////////////////////////////////////////////////

Implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  CRevPlan_TOM, // pour ne pas l'oublier dans les projets
  FamRub_Tof,   // ParametragePlanRevision
  Ent1,         // EstSpecif
  uLibRevision, // TAligneRevision
  Choix, Mask;  // Choisir

const cColGene   = 0;
      cColExclu  = 1;
      cColActive = 2;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CREVPARAMCYCLE( vStLequel : string; vStActionFiche : TActionFiche ) : string;
begin
  if not ExisteSQL('SELECT CPR_PLANREVISION FROM CREVPLAN') then
    PgiInfo('Vous devez d''abord paramétrer les plans de révision.', 'Cycle de révision')
  else
    Result := AGLLanceFiche('CP', 'CPREVPARAMCYCLE', '', vStLequel, ActionToString(vStActionFiche));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnArgument ( S: String ) ;
begin
  Inherited ;
  Pages         := THPageControl2(GetControl('PAGES', True));
  FListe        := THGrid(GetControl('FLISTE', True));
  ListePlan     := TListBox(GetControl('LISTEPLAN', True));

  BPlanRevision := TToolbarButton97(GetControl('BPLANREVISION', True));
  BCopier       := TToolbarButton97(GetControl('BCOPIER', True));
  BEfface       := TToolbarButton97(GetControl('BEFFACE', True));
  BDelete       := TToolBarbutton97(GetControl('BDELETE', True));

  Z_C1          := THValComboBox(GetControl('Z_C1', True));
  ZO1           := THValComboBox(GetControl('ZO1', True));
  ZV1           := THEdit(GetControl('ZV1', True));
  ZG1           := THValComboBox(GetControl('ZG1', True));

  Z_C2          := THValComboBox(GetControl('Z_C2', True));
  ZO2           := THValComboBox(GetControl('ZO2', True));
  ZV2           := THEdit(GetControl('ZV2', True));
  ZG2           := THValComboBox(GetControl('ZG2', True));

  Z_C3          := THValComboBox(GetControl('Z_C3', True));
  ZO3           := THValComboBox(GetControl('ZO3', True));
  ZV3           := THEdit(GetControl('ZV3', True));

{$IFDEF EAGLCLIENT}
  THEdit(GetControl('CPC_CODECYCLE', True)).OnKeyPress := OnKeyPressCPC_CODECYCLE;
{$ELSE}
  THDBEdit(GetControl('CPC_CODECYCLE', True)).OnKeyPress := OnKeyPressCPC_CODECYCLE;
{$ENDIF}

  // Branchement des événements
  FListe.OnDblClick     := OnDblClickFListe;
  FListe.OnCellExit     := OnCellExitFListe;
  FListe.OnKeyPress     := OnKeyPressFListe;
  FListe.OnKeyDown      := OnKeyDownFListe;

  BPlanRevision.OnClick := OnClickBPlanRevision;
  BCopier.OnClick       := OnClickBCopier;
  BEfface.OnClick       := OnClickBEfface;

  Z_C1.OnChange         := OnChangeZC;
  ZO1.OnChange          := OnChangeZxx;
  ZV1.OnChange          := OnChangeZxx;

  //fOnSaveElipsisClick := ZV1.OnClick;
  //ZV1.onElipsisClick  := OnElipsisClickZV;
  ZG1.OnChange          := OnChangeZxx;

  Z_C2.OnChange         := OnChangeZC;
  ZO2.OnChange          := OnChangeZxx;
  ZV2.OnChange          := OnChangeZxx;
  //ZV2.OnElipsisClick    := OnElipsisClickZV;
  ZG2.OnChange          := OnChangeZxx;

  Z_C3.OnChange         := OnChangeZC;
  ZO3.OnChange          := OnChangeZxx;
  ZV3.OnChange          := OnChangeZxx;
  //ZV3.OnElipsisClick    := OnElipsisClickZV;

  FListe.ColAligns [cColGene]   := TaLeftJustify;
  FListe.ColAligns [cColExclu]  := TaLeftJustify;
  FListe.ColAligns [cColActive] := TaCenter;
  FListe.ColTypes  [cColActive] := 'B';
  FListe.ColFormats[cColActive] := IntToStr( Byte( csCheckBox ) );

  // GCO - 09/07/2007 - FQ 20890
  SetControlProperty('CPC_CODECYCLE', 'MAXLENGTH', 5);

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnClose ;
begin
  Inherited ;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnLoadRecord ;
begin
  Inherited ;
  FBoLoading := True;
  ChargeEnregistrement;
  BCopier.Enabled       := not (TFFiche(Ecran).FTypeAction in [TaConsult]);
  BPlanRevision.Enabled := not (TFFiche(Ecran).FTypeAction in [TaConsult]);
  // GCO - 11/07/2007 - FQ 20883
  BEfface.Visible       := not (TFFiche(Ecran).FTypeAction in [TaConsult]);

  // GCO - 06/06/2007 - FQ 20566 + FQ 20883
{$IFDEF CCSTD}
  BDelete.Visible := (((GetField('CPC_PREDEFINI') = 'CEG') and EstSpecif('51502')) or
                       (GetField('CPC_PREDEFINI') = 'STD')) and
                     (not (TFFiche(Ecran).FTypeAction in [TaConsult]));
{$ELSE}
  BDelete.Visible := False;
{$ENDIF}

  TTabSheet(GetControl('TABACTIVATION', True)).TabVisible := VH^.OkModRic and (CtxPcl in V_Pgi.PgiContexte);

  if FListe.CanFocus then
    FListe.SetFocus;

  FBoLoading := False;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.ChargeEnregistrement;
begin
  FStPlanAssocie    := GetField('CPC_PLANASSOCIE');
  FStListeCompte    := GetField('CPC_LISTECOMPTE');
  FStListeExclusion := GetField('CPC_LISTEEXCLUSION');

  ChargeListePlan;
  ChargeFListe;
  ChargeActivation;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.ChargeListePlan;
var lStLibPlanRevision : string;
    lStringListPlan : TStringList;
    lQuery : TQuery;
    i : integer;
    lSt, lSt1 : string;
begin
  lStLibPlanRevision := GetField('CPC_PLANASSOCIE');
  ListePlan.Items.Clear ;
  lStringListPlan := TStringList.Create;

  lQuery := nil;
  try
    lQuery := OpenSQL('SELECT CPR_PLANREVISION, CPR_LIBELLE FROM CREVPLAN ' +
                      'ORDER BY CPR_PLANREVISION', True);
    while not lQuery.Eof do
    begin
      lStringListPlan.Add(lQuery.FindField('CPR_PLANREVISION').AsString + '=' +
                          lQuery.FindField('CPR_LIBELLE').AsString);
      lQuery.Next;
    end;
  finally
    Ferme(lQuery);
  end;

  while lStLibPlanRevision <> '' do
  begin
    lSt := ReadTokenSt(lStLibPlanRevision);
    i   := lStringListPlan.IndexOfName(lSt);
    if i >= 0 then
      lSt1 := lStringListPlan.Values[lSt]
    else
      lSt1 := 'Non renseigné' ;

    ListePlan.Items.Add(lSt1) ;
  end;

  lStringListPlan.Free;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.ChargeFListe;
var lStListeGene   : string;
    lStListeExclu  : string;
    lStListeActive : string;
    lStGene        : string;
    lStExclu       : string;
    lStActive      : string;

    lRow           : integer;
begin
  try
    lRow := 0;
    lStListeGene   := GetField('CPC_LISTECOMPTE');
    lStListeExclu  := GetField('CPC_LISTEEXCLUSION');
    lStListeActive := GetField('CPC_LISTEACTIVE');

    while lStListeGene <> '' do
    begin
      lStGene   := ReadTokenSt(lStListeGene);
      lStExclu  := ReadTokenSt(lStListeExclu);
      lStActive := ReadTokenSt(lStListeActive);

      if lStGene <> '' then
      begin
        if (lRow <> 0 ) then FListe.RowCount := FListe.RowCount+1;
        Inc(lRow);
        FListe.Cells[cColGene,lRow]   := lStGene;
        FListe.Cells[cColExclu,lRow]  := lStExclu;
        FListe.Cells[cColActive,lRow] := lStActive;
      end;
    end;

  finally
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.ChargeActivation;
var lSt : string;
    lStTemp : string;
begin
  lSt := GetField('CPC_ACTIVATION');

  lSt := FindEtReplace(lSt, '(', '', True);
  lSt := FindEtReplace(lSt, ')', '', True);

  lStTemp      := ReadTokenPipe(lSt, '|');
  Z_C1.Value   := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZO1.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZV1.Text     := lStTemp;
  ZV1.DataType := TrouveDataType( Z_C1.Value );

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZG1.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  Z_C2.Value   := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZO2.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZV2.Text     := lStTemp;
  ZV2.DataType := TrouveDataType( Z_C2.Value );

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZG2.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  Z_C3.Value   := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZO3.Value    := lStTemp;

  lStTemp      := ReadTokenPipe(lSt, '|');
  ZV3.Text     := lStTemp;
  ZV3.DataType := TrouveDataType( Z_C3.Value );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... : 23/04/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnNewRecord ;
begin
  Inherited ;
  if ctxStandard in V_PGI.PGIContexte then
  begin
    if EstSpecif('51502') and (ctxStandard in V_PGI.PGIContexte) then
      SetField('CPC_PREDEFINI', 'CEG')
    else
      SetField('CPC_PREDEFINI', 'STD');
  end
  else
    SetField('CPC_PREDEFINI', 'STD');

  FListe.VidePile(False);
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnUpdateRecord;
var lRic : TRic;
    lStCycleDoublon : String;
begin
  Inherited ;

  // GCO - 28/09/2007 - FQ 21534
  if GetField('CPC_PLANASSOCIE') = '' then
  begin
    PgiInfo('Vous devez associer un plan de révision au cycle.', 'Cycle de révision');
    LastError := -1;
    Exit;
  end;

  SauveFListe;
  SauveActivation;

  lStCycleDoublon := '';
  lRic := nil;
  try
    lRic := TRic.Create;
    if not lRic.ControleUnicite( GetField('CPC_CODECYCLE'),
      GetField('CPC_PREDEFINI'), GetField('CPC_LISTECOMPTE'),
      GetField('CPC_LISTEEXCLUSION'), GetField('CPC_PLANASSOCIE'), lStCycleDoublon) then
    begin
      // Message d'information si le cycle a des condtiions d'activation
      if GetField('CPC_ACTIVATIONSQL') <> '' then
        PgiInfo('A titre d''information. Erreur de fourchette de comptes avec le cycle ' +
                lStCycleDoublon + '.', 'Contrôle d''unicité')
      else // Message bloquant si aucune présence de conditions d'activation
      begin
        PgiInfo('Enregistrement annulé. Erreur de fourchette de comptes avec le cycle ' + #13#10 +
                'de révision ' + lStCycleDoublon + '.', 'Contrôle d''unicité');
        LastError := -1;
        Exit;
      end;
    end;
  finally
    FreeAndNil( lRic );
  end;

  // GCO - 03/10/2007 - FQ 21564
  if (FStPlanAssocie <> GetField('CPC_PLANASSOCIE')) or
     (FStListeCompte <> GetField('CPC_LISTECOMPTE')) or
     (FStListeExclusion <> GetField('CPC_LISTEEXCLUSION')) then
  begin
    if not MAJDateModifCREVPLAN then
    begin
      LastError := -1;
      Exit;
    end;
  end;
  // FIN GCO

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnDeleteRecord ;
begin
  Inherited ;

  if not MAJDateModifCRevPlan then
    LastError := -1;
    
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnCancelRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnClickBPlanRevision(Sender: TObject);
var lStPlanRevision    : string;
    lStLibPlanRevision : string;
begin
  if ListePlan.CanFocus then
    ListePlan.SetFocus;

  // GCO - 31/05/2007 - FQ 20490
  if not ExisteSQL('SELECT CPR_PLANREVISION FROM CREVPLAN WHERE ' +
                   'CPR_PREDEFINI = "' + GetField('CPC_PREDEFINI') +  '"') then
  begin
    PgiInfo('Vous devez d''abord créer un plan de révision personalisé.', Ecran.Caption);
    Exit ;
  end;                 

  if Trim(GetControlText('CPC_CODECYCLE')) = '' then
  begin
    PgiInfo('Vous devez renseigner un code cycle.', Ecran.Caption);
    Exit ;
  end;

  lStPlanRevision := GetField('CPC_PLANASSOCIE');

  ParametragePlanRevision(GetControlText('CPC_CODECYCLE'),
                          GetControlText('CPC_LIBELLECYCLE'),
                          lStPlanRevision, lStLibPlanRevision);

  if lStPlanRevision <> GetField('CPC_PLANASSOCIE') then
  begin
    if DS.State = dsBrowse then ForceUpdate;
    if lStPlanRevision <> '' then
      SetField('CPC_PLANASSOCIE', lStPlanRevision);

    ListePlan.Items.Clear ;
    while lStLibPlanRevision <> '' do
      ListePlan.Items.Add(ReadTokenSt(lStLibPlanRevision));
  end ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnClickBCopier(Sender: TObject);
var lSt : string;
    lStWhere : string;
    lQuery : TQuery;
begin
  if Trim(GetControlText('CPC_CODECYCLE')) = '' then
  begin
    PgiInfo('Vous devez renseigné le code identifiant du cycle de révision.', Ecran.Caption);
    SetFocusControl('CPC_CODECYCLE');
    Exit ;
  end;

  // GCO - 24/07/2007 - FQ 21164
  //lStWhere := 'CPC_CODECYCLE <> "' + GetControlText('CPC_CODECYCLE') + '"';

  // Choix du cycle de révision à dupliquer
  lSt := Choisir(TraduireMemoire('Choix d''un cycle de révision à dupliquer'),
                 'CREVPARAMCYCLE', 'CPC_CODECYCLE || " / " || CPC_LIBELLECYCLE',
                 'CPC_CODECYCLE', lStWhere, 'CPC_CODECYCLE');

  if lSt = '' then Exit;

  // Confirmation de la duplication
  if PgiAsk('Confirmez-vous la duplication du cycle de révision sélectionné ?', Ecran.Caption) <> mrYes then
    Exit;

  // Changement l'état si nécessaire
  if not (TFFiche(Ecran).FTypeAction in [taCreat, taCreatOne, taCreatEnSerie]) then
    TFFiche(Ecran).FTypeAction := taCreat;

  lQuery := OpenSQL('SELECT * FROM CREVPARAMCYCLE WHERE ' +
                    'CPC_CODECYCLE = "' + lSt + '"', True);

  if not lQuery.Eof then
  begin
    if Trim(GetField('CPC_CODEABREGE')) = '' then
      SetField('CPC_CODEABREGE', Copy(GetControlText('CPC_CODECYCLE'), 1, 6));

    SetField('CPC_LIBELLECYCLE',   lQuery.FindField('CPC_LIBELLECYCLE').AsString);

    // GCO - 24/07/2007 - FQ 21166
    // Solution logique, on oblige l'utilisateur à choisir un plan autre que
    // celui de départ.
    //SetField('CPC_PLANASSOCIE',    lQuery.FindField('CPC_PLANASSOCIE').AsString);

    SetField('CPC_LISTECOMPTE',    lQuery.FindField('CPC_LISTECOMPTE').AsString);
    SetField('CPC_LISTEEXCLUSION', lQuery.FindField('CPC_LISTEEXCLUSION').AsString);
    SetField('CPC_LISTEACTIVE',    lQuery.FindField('CPC_LISTEACTIVE').AsString);
    SetField('CPC_ACTIVATION',     lQuery.FindField('CPC_ACTIVATION').AsString);
    SetField('CPC_ACTIVATIONSQL',  lQuery.FindField('CPC_ACTIVATIONSQL').AsString);

    if ctxStandard in V_PGI.PGIContexte then
    begin
      if EstSpecif('51502') and (ctxStandard in V_PGI.PGIContexte) then
        SetField('CPC_PREDEFINI', 'CEG')
      else
        SetField('CPC_PREDEFINI', 'STD');
    end
    else
      SetField('CPC_PREDEFINI', 'STD');
  end ;

  Ferme(lQuery);

  ChargeEnregistrement;
  SetControlEnabled('CPC_CODECYCLE', True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/03/2007
Modifié le ... : 05/03/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnClickBEfface(Sender: TObject);
begin
  Z_C1.ItemIndex := -1;
  ZO1.ItemIndex := -1;
  ZV1.Text := '';
  ZG1.ItemIndex := -1;

  Z_C2.ItemIndex := -1;
  ZO2.ItemIndex := -1;
  ZV2.Text := '';
  ZG2.ItemIndex := -1;

  Z_C3.ItemIndex := -1;
  ZO3.ItemIndex := -1;
  ZV3.Text := '';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.SauveFListe;
var i : integer;
    lStListeGene   : string;
    lStListeExclu  : string;
    lStListeActive : string;
begin
  lStListeGene   := '';
  lStListeExclu  := '';
  lStListeActive := '';

  for i := 1 to FListe.RowCount -1 do
  begin
    if Trim(FListe.Cells[cColGene,i]) <> '' then
    begin
      lStListeGene   := lStListeGene   + FListe.Cells[cColGene,i]   + ';';
      lStListeExclu  := lStListeExclu  + FListe.Cells[cColExclu,i]  + ';';
      lStListeActive := lStListeActive + FListe.Cells[cColActive,i] + ';';
    end;  
  end;

  SetField('CPC_LISTECOMPTE',    lStListeGene);
  SetField('CPC_LISTEEXCLUSION', lStListeExclu);
  SetField('CPC_LISTEACTIVE',    lStListeActive);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.SauveActivation;
var lSt : string;
begin
  lSt := '';

  if (Z_C1.Value <> '') and (ZO1.Value <> '') and (ZV1.Text <> '') then
  begin
    lSt := lSt + Z_C1.Value + '|' + ZO1.Value + '|' + ZV1.Text + '|' ;

    if (ZG1.Value <> '') and (Z_C2.Value <> '') and
       (ZO2.Value <> '') and (ZV2.Text <> '') then
    begin
      lSt := lSt + ZG1.Value + '|' + Z_C2.Value + '|' + ZO2.Value + '|' + ZV2.Text + '|';

      if (ZG2.Value <> '') and (Z_C3.Value <> '') and
         (ZO3.Value <> '') and (ZV3.Text <> '') then
        lSt := lSt + ZG2.Value + '|' + Z_C3.Value + '|' + ZO3.Value + '|' + ZV3.Text;
    end;
  end;

  SetField('CPC_ACTIVATION', lSt);
  lSt := RecupWhereCritere( Pages );
  SetField('CPC_ACTIVATIONSQL', Copy(lSt, 7, Length(lSt)));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnCellExitFListe(Sender : TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if TFFiche(Ecran).FTypeAction = taConsult then Exit;

  if (ACol = cColActive) and ( ARow = FListe.RowCount -1) then
  begin
    if FListe.Cells[cColGene,ARow] <> '' then
    begin
      if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;
      FListe.RowCount := FListe.RowCount + 1;
      FListe.Row := FListe.RowCount - 1;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnKeyDownFListe(Sender : TObject; var Key: Word; Shift: TShiftState);
begin
  if TFFiche(Ecran).FTypeAction = TaConsult then Exit;
  if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnKeyPressFListe(Sender : TObject; var Key: Char);
begin
  if TFFiche(Ecran).FTypeAction = TaConsult then Exit;
  if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;

  if FListe.Col = cColActive then
  begin
    if not (Key in ['X','-','x']) then
      Key := #0
    else
      Key := UpCase(Key);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnDblClickFListe(Sender: TObject);
begin
  if TFFiche(Ecran).FTypeAction = TaConsult then Exit;
  if (FListe.Row > 0) and (FListe.Col = cColActive) then
  begin
    if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;

    if FListe.Cells[FListe.Col, FListe.Row] = 'X' then
      FListe.Cells[FListe.Col, FListe.Row] := '-'
    else
      FListe.Cells[FListe.Col, FListe.Row] := 'X';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnChangeZC(Sender: TObject);
var lSt  : string;
begin
  if FBoLoading then Exit;
  if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;

  lSt := TrouveDataType( THValComBoBox(Sender).Value );

  if THValComBoBox(Sender).Name = 'Z_C1' then
  begin
    ZV1.Text := '';
    ZV1.DataType  := lSt;
  end;

  if THValComBoBox(Sender).Name = 'Z_C2' then
  begin
    ZV2.Text := '';
    ZV2.DataType  := lSt;
  end;

  if THValComBoBox(Sender).Name = 'Z_C3' then
  begin
    ZV3.Text := '';
    ZV3.DataType  := lSt;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnChangeZxx(Sender: TObject);
begin
  if FBoLoading then Exit;
  if not (DS.State in [dsEdit,dsInsert]) then
    DS.Edit;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnElipsisClickZV(Sender: TObject);
begin
  //if FBoLoading then Exit;
  //if not (DS.State in [dsEdit,dsInsert]) then DS.Edit ;
  //if Assigned(fOnSaveElipsisClick) then
  //  fOnSaveElipsisClick( Sender );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_CREVPARAMCYCLE.TrouveDataType(vStValue : string) : string;
begin
  Result := '';
  if vStValue = '' then Exit;
  Result := GetColonneSQL('COMMUN', 'CO_LIBRE', 'CO_TYPE = "CDA" AND ' +
                          'CO_ABREGE = "' + vStValue + '"');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVPARAMCYCLE.OnKeyPressCPC_CODECYCLE(Sender: TObject; var Key: Char);
begin
  // GCO - 18/09/2007 - FQ 21436
  // Autorise uniquement les caractères A-Z et 0-9 et Backspace
 if not(Key in ['a'..'z','A'..'Z','0'..'9',#8]) then Key := #0;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_CREVPARAMCYCLE.MAJDateModifCREVPLAN : boolean;
var lStPlan         : HString;
    lStwherePlan    : HString;
begin
  Result := True;
  try
      lStPlan := GetField('CPC_PLANASSOCIE') + FStPlanAssocie;
      lStWherePlan := '(';
      while Pos(';', lStPlan) > 0 do
      begin
        lStWherePlan := lStWherePlan + ' CPR_PLANREVISION = "' + ReadTokenSt( lStPlan ) + '" OR ';
      end;
      // Suppression du ' OR '
      lStWherePlan := Copy(lStWherePlan, 0, Length(lStWherePlan) - 3) + ')';

      // GCO - 16/11/2007 - FQ 21869
      ExecuteSQL('UPDATE CREVPLAN SET CPR_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' +
                 lStWherePlan);
  except
    on E: Exception do
    begin
      Result := False;
      PgiError('Erreur lors de la mise à jour de la date de modification des plans de révision'+
                #13+#10 +  E.Message, 'TOM_CREVPARAMCYCLE.OnUpdateRecord');
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOM_CREVPARAMCYCLE ] ) ;
end.

