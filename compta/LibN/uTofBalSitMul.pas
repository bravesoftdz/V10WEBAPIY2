{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : BALSIT_MUL ()
Mots clefs ... : TOF;BALSIT_MUL
*****************************************************************}
Unit uTofBalSitMul ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils,  ComCtrls,
     Menus, HCtrls, HEnt1, HMsgBox, UTOF, HTB97, HDB, BalSit, uTob, HQry,
     Paramsoc, // GetParamSocSecur

{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     fe_main,
     mul,
{$ENDIF}
     CritEdt,         // ttypEcr
     Ent1;


Type
  TOF_BALSIT_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
{$IFDEF EAGLCLIENT}
      fListe : THGrid;
{$ELSE}
      fListe : THDBGrid;
{$ENDIF}
      FPopZoom : TPopupMenu;
      procedure OnSupprimerClick ( Sender : TOBject );
      procedure OnInsertClick    ( Sender : TOBject );
      procedure OnImporterClick  ( Sender : TOBject );
      procedure OnOuvrirClick    ( Sender : TOBject );
      procedure OnCalculClick    ( Sender : TOBject );
      procedure OnFigerClick     ( Sender : TOBject );

      procedure SupprimeBalance;
      procedure FigeBalSit (pstCode : string);
      procedure OnDblClickListe ( Sender : TObject );

      procedure ReCalculBalanceSituation( vStCodeBal : string );  
      procedure OnKeyDownEcran(Sender : TObject ; var Key : Word ; Shift : TShiftState); // FQ 21518 BVE 26.09.07
  end ;

procedure CCLanceFiche_BalanceSituation;

Implementation

uses AGLInit,            // TheMulQ
     Windows,            // VK_DELETE FQ 21518 
     uTofBalsitModif,    // CPLanceFiche_BALSIT_MODIF
     uTofBalsitCreation, // CPLanceFiche_BALSIT_CREATION
     uTofBalsitImport;   // CPLanceFiche_BALSIT_IMPORT

procedure CCLanceFiche_BalanceSituation;
begin
  AGLLanceFiche ('CP','BALSIT_MUL','','','');
end;

procedure TOF_BALSIT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_MUL.OnArgument (S : String ) ;
begin
  Inherited ;

  Ecran.HelpContext:=999999107;

  FPopZoom := TPopupMenu (GetControl ('MONPOPUP'));
  AddMenuPop (FPopZoom,'','');
  TToolBarButton97(GetControl('BSUPPRIMER')).OnClick := OnSupprimerClick ;
  TToolBarButton97(GetControl('BINSERT')).OnClick    := OnInsertClick ;
  TToolBarButton97(GetControl('BOUVRIR')).OnClick    := OnOuvrirClick ;
  TToolBarButton97(GetControl('BIMPORTER')).OnClick  := OnImporterClick ;
  TToolBarButton97(GetControl('BCALCUL')).OnClick    := OnCalculClick;
  TToolBarButton97(GetControl('BFIGER')).OnClick    := OnFigerClick;
  (*
  if EstComptaSansAna then begin
    SetControlProperty('BSI_TYPECUM', 'PLUS', ' AND CO_CODE<>"G/A" AND CO_CODE<>"ANA"');
    SetControlVisible('BSI_AXE', False);
    SetControlVisible('TBSI_AXE', False);
  end;
  *)
{$IFDEF EAGLCLIENT}
  fListe := THGrid (GetControl ('FLISTE'));
{$ELSE}
  fListe := THDBGrid (GetControl ('FLISTE'));
{$ENDIF}
  fListe.OnDblClick := OnDblClickListe;

  PositionneEtabUser(GetControl('BSI_ETABLISSEMENT'), False); // 15090

  THValComboBox(GetControl('BSI_TYPECUM')).ItemIndex := 0;
  THValComboBox(GetControl('BSI_TYPEBAL')).ItemIndex := 0;

  if not GetParamSocSecur('SO_CPBDSDYNA', False) then
  begin
    SetControlVisible('BSI_TYPEBAL', False);
    SetControlVisible('TBSI_TYPEBAL', False);
    SetControlVisible('BCALCUL', False);
  end;
  { FQ 21518 BVE 26.09.07 }
  Ecran.OnKeyDown := OnKeyDownEcran;
  { END FQ 21518 }
end ;

procedure TOF_BALSIT_MUL.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/07/2005
Modifié le ... : 20/07/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_BALSIT_MUL.OnInsertClick(Sender: TOBject);
begin
  CPLanceFiche_BALSIT_CREATION;
  TFMUL(ECRAN).BChercheClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_BALSIT_MUL.OnOuvrirClick(Sender: TOBject);
var lQuery : TQuery;
begin
{$IFDEF EAGLCLIENT}
  lQuery := TFMul(Ecran).Q.TQ;
{$ELSE}
  lQuery := TFMul(Ecran).Q;
{$ENDIF}
  if not lQuery.Eof then
  begin
    CPLanceFiche_BALSIT_MODIF( GetField('BSI_CODEBAL'));
    TFMUL(ECRAN).BChercheClick(nil);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_BALSIT_MUL.OnCalculClick(Sender: TOBject);
var lQuery  : TQuery;
    i : integer;
begin
{$IFDEF EAGLCLIENT}
  lQuery := TFMul(Ecran).Q.TQ;
{$ELSE}
  lQuery := TFMul(Ecran).Q;
{$ENDIF}

  if (FListe.NbSelected = 0 ) and (not FListe.AllSelected) then
  begin
    PGIInfo ('Aucun enregistrement sélectionné.',ECRAN.Caption);
    Exit;
  end;

  if PGIAsk ('Êtes-vous sûr de vouloir recalculer les balances de situation dynamiques sélectionnées ?', ECRAN.Caption) <> mrYes then
  begin
    FListe.ClearSelected;
    Exit;
  end;

  if FListe.AllSelected then
  begin
    try
      BeginTrans;
      lQuery.First;
      while not lQuery.Eof do
      begin
        ReCalculBalanceSituation(lQuery.FindField('BSI_CODEBAL').AsString);
        lQuery.Next;
      end;
      CommitTrans;
    except
      RollBack;
    end;
  end
  else
  begin
    try
      BeginTrans;
      for i := 0 to FListe.nbSelected - 1 do
      begin
        FListe.GotoLeBookmark(i);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
        {$ENDIF}
        ReCalculBalanceSituation(GetField('BSI_CODEBAL'));
      end;
      CommitTrans;
    except
      RollBack;
    end;
  end;

  FListe.ClearSelected;
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_BALSIT_MUL.OnImporterClick(Sender: TOBject);
begin
  CPLanceFiche_BALSIT_IMPORT;
  TFMUL(ECRAN).BChercheClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_BALSIT_MUL.OnSupprimerClick(Sender: TOBject);
var i : integer;
    bErreur : boolean;
    stWhere : string;
    TBal : TOB;
    BalSit : TBalSit;
    bPresenceBDSDynamique : boolean;
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

  if (fListe.NbSelected = 0 ) and (not fListe.AllSelected)  then
  begin
    PGIInfo ('Aucun enregistrement sélectionné.',ECRAN.Caption);
    exit;
  end;
  if PGIAsk ('Êtes-vous sûr de vouloir supprimer les enregistrements sélectionnés ?',ECRAN.Caption) <> mrYes then
  begin
    fListe.ClearSelected;
    exit;
  end;

  bPresenceBDSDynamique := False;

  if fListe.AllSelected then
  begin
    stWhere := RecupWhereCritere(TFMul(Ecran).Pages);
    TBal := TOB.Create ('', nil, -1);
    TBal.LoadDetailFromSQL('SELECT * FROM CBALSIT WHERE '+stWhere);
    for i:=0 to TBal.Detail.Count - 1 do
    begin
      // GCO - 03/07/2006 - Interdir la suppression des BDS DYN CEGID : BDS@xxxx
      if (Copy(TBal.Detail[i].GetString('BSI_CODEBAL'), 1, 4) = 'BDS@') then
      begin
        bPresenceBDSDynamique := True;
        Continue;
      end;

      BeginTrans;
      try
        BalSit := TBalSit.Create (TBal.Detail[i].GetValue('BSI_CODEBAL'));
        BalSit.Supprime;
        BalSit.Free;
        CommitTrans;
      except
        RollBack;
        MessageAlerte('Suppression de la balance '+TBal.Detail[i].GetValue('BSI_CODEBAL')+' impossible.');
      end;
    end;
    TBal.Free;
  end
  else
  begin
    for i:=0 to fListe.nbSelected - 1 do
    begin
      fListe.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
      {$ENDIF}
      if (Copy(TheMulQ.FindField('BSI_CODEBAL').AsString, 1, 4) = 'BDS@') then
      begin
        bPresenceBDSDynamique := True;
        Continue;
      end;
      if ( Transactions ( SupprimeBalance , 1 ) <> oeOK ) then
      begin
        bErreur := True;
        MessageAlerte('Suppression de la balance '+ TheMulQ.FindField('BSI_LIBELLE').AsString+' impossible.');
      end else bErreur := False;
      if bErreur then break;
    end;
  end;

  if BPresenceBDSDynamique then
    PgiInfo('Les balances de situation dynamiques CEGID n''ont pas été supprimées.', Ecran.Caption)
  else
    PGIInfo ('Suppression terminée.',ECRAN.Caption);

  fListe.ClearSelected;
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

procedure TOF_BALSIT_MUL.SupprimeBalance;
var BalSit : TBalSit;
begin
  BalSit := TBalSit.Create (GetField ('BSI_CODEBAL'));
  BalSit.Supprime;
  BalSit.Free;
end;

procedure TOF_BALSIT_MUL.OnDblClickListe(Sender: TObject);
begin
  OnOuvrirClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_BALSIT_MUL.ReCalculBalanceSituation(vStCodeBal: string);
var lInfoBS : RBalSitInfo;
    lBS : TBalSit;
    lTobBalSit : Tob;
begin
  lTobBalSit := Tob.Create('CBALSIT', nil, -1);
  lTobBalSit.SelectDB('"'+ vStCodeBal + '"', nil);

  // On ne recalcul que les balances dynamiques
  if lTobBalSit.GetString('BSI_TYPEBAL') = 'DYN' then
  begin
    ExecuteSQL('DELETE FROM CBALSITECR WHERE BSE_CODEBAL = "' + vStCodeBal + '"');
    ExecuteSQL('DELETE FROM CBALSIT WHERE BSI_CODEBAL = "' + vStCodeBal+ '"');

    lBS := TBalSit.Create(vStCodeBal);

    lInfoBS.Libelle        := lTobBalSit.GetString('BSI_LIBELLE');
    lInfoBS.Abrege         := lTobBalSit.GetString('BSI_ABREGE');
    lInfoBS.Plan           := lTobBalSit.GetString('BSI_TYPECUM');
    lInfoBS.DateInf        := lTobBalSit.GetDateTime('BSI_DATE1');
    lInfoBS.DateSup        := lTobBalSit.GetDateTime('BSI_DATE2');
    lInfoBS.Axe            := lTobBalSit.GetString('BSI_AXE');
    lInfoBS.Exo            := lTobBalSit.GetString('BSI_EXERCICE');

    if lTobBalSit.GetString('BSI_ETABLISSEMENT') <> '...' then
      lInfoBS.Etab           := lTobBalSit.GetString('BSI_ETABLISSEMENT');

    lInfoBS.Devise         := lTobBalSit.GetString('BSI_DEVISE');

    lInfoBs.Ano            := True;
    lInfoBS.TypeBalance    := 'DYN';

    lInfoBS.TypeEcr[Reel]  := True;
    lInfoBS.TypeEcr[Simu]  := False;
    lInfoBS.TypeEcr[Situ]  := False;
    lInfoBS.TypeEcr[Previ] := False;
    lInfoBS.TypeEcr[Revi]  := False;
    lInfoBS.TypeEcr[Ifrs]  := False;

    lInfoBS.Compte1Inf     := '';
    lInfoBS.Compte1Sup     := '';
    lInfoBS.Compte2Inf     := '';
    lInfoBS.Compte2Sup     := '';

    lInfoBS.CumExo         := False;
    lInfoBS.CumMois        := False;
    lInfoBS.CumPer         := False;
    lInfoBS.CumEtab        := False;
    lInfoBS.CumDev         := False;
    lInfoBS.DebCreColl     := False;

    lBS.GenereAuto ( lInfoBS );
    lBS.Free;
  end;
  FreeAndNil(lTobBalSit);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_BALSIT_MUL.OnFigerClick(Sender: TOBject);
var lQuery  : TQuery;
    i : integer;
begin
{$IFDEF EAGLCLIENT}
  lQuery := TFMul(Ecran).Q.TQ;
{$ELSE}
  lQuery := TFMul(Ecran).Q;
{$ENDIF}

  if (FListe.NbSelected = 0 ) and (not FListe.AllSelected) then
  begin
    PGIInfo ('Aucun enregistrement sélectionné.',ECRAN.Caption);
    Exit;
  end;

  if PGIAsk ('Voulez-vous figer les balances de situation dynamiques sélectionnées ?', ECRAN.Caption) <> mrYes then
  begin
    FListe.ClearSelected;
    Exit;
  end;

  if FListe.AllSelected then
  begin
    try
      BeginTrans;
      lQuery.First;
      while not lQuery.Eof do
      begin
        FigeBalSit(lQuery.FindField('BSI_CODEBAL').AsString);
        lQuery.Next;
      end;
      CommitTrans;
    except
      RollBack;
    end;
  end
  else
  begin
    try
      BeginTrans;
      for i := 0 to FListe.nbSelected - 1 do
      begin
        FListe.GotoLeBookmark(i);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
        {$ENDIF}
        FigeBalsit(GetField('BSI_CODEBAL'));
      end;
      CommitTrans;
    except
      RollBack;
    end;
  end;

  FListe.ClearSelected;
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

procedure TOF_BALSIT_MUL.FigeBalSit(pstCode: string);
begin
 ExecuteSQL ('UPDATE CBALSIT SET BSI_TYPEBAL="BDS" WHERE BSI_CODEBAL="'+pstCode+'" AND BSI_TYPEBAL="DYN"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 26/09/2007
Modifié le ... :   /  /    
Description .. : Rajout du CTRL + SUPPR dans la liste
Suite ........ : FQ 21518
Mots clefs ... : 
*****************************************************************}
procedure TOF_BALSIT_MUL.OnKeyDownEcran(Sender : TObject ; var Key : Word ; Shift : TShiftState);
begin
   case key of
      VK_DELETE :
      begin         
         if FListe.Focused and (Shift = [ssCtrl]) then
         begin
             OnSupprimerClick (Sender);
             Key := 0;
         end;
      end;
   end;
   if ( Ecran <> nil ) and ( Ecran is  TFMul ) then
      TFMul(Ecran).FormKeyDown(Sender,Key,Shift);
end;

Initialization
  registerclasses ( [ TOF_BALSIT_MUL ] ) ;
end.
