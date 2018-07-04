{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/03/2001
Modifié le ... : 16/03/2001
Description .. : Source TOF de la TABLE : ICCGENERAUX ()
Mots clefs ... : TOF;ICCGENERAUX
*****************************************************************}
unit uTOFIccGeneraux;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  UTOb,
  MaineAGL,
  UtileAGL,
  eMul,    // TFMUL(ECRAN).BChercheClick( nil )
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main, // AGLLanceFiche
  Hdb,
  Mul,     // TFMUL(ECRAN).BChercheClick( nil )
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTof,
  Windows,
  ParamSoc,
  Menus,
  Htb97,
  Ent1,
  uTom,    // GetField
  AGLInit, // TheMulQ
  uTot;

procedure CPLanceFiche_ICCGENERAUX;

type
  TOF_ICCGENERAUX = class(TOF)
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;

    procedure OnDblClickFListe( Sender : TObject );

  private
    MonPopUp, PopUpImporter: TPopUpMenu;
    BInsert, BDelete, BCalculer: TToolBarButton97;

{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}

    Param       : string;
    FDateCtxDeb : TDateTime;
    FDateCtxFin : TDateTime;

    procedure OnKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnClickBInsert(Sender: TObject);
    procedure OnClickBDelete(Sender: TObject);
    procedure OnClickBCalculer(Sender: TObject);
    procedure OnClickBRecuperation(Sender: TObject);
    procedure OnClickBMiseAJour(Sender: TObject);
    procedure SuppressionCompte;
    procedure ImportationEcriture(ModeMiseAjour: Boolean);
    procedure RechargeTheMulQ;

  end;

implementation

uses uLibExercice,     // CExoRefOuvert
     uLibWindows,      // IIF(
     IccGlobale,       // Icc_Data, TestImpositionSociete
     uTOMIccGeneraux,  // CPLanceFiche_CreateICCGENERAUX
     uTOFIccParametre; // CPLanceFiche_ICCPARAMETRE

////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_ICCGENERAUX;
begin
  AGLLanceFiche('CP', 'ICCGENERAUX', '', '', 'ACTION=CONSULTATION');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_ICCGENERAUX.OnArgument(S: string);
var StArg: string;
begin
  inherited;

  Ecran.HelpContext:=999999102;

  StArg := S;

  Param := ReadTokenST(StArg); // Lecture de ACTION
{$IFNDEF EAGLCLIENT}
  THDBGrid(GetControl('FLISTE')).OnKeyUp := OnKeyUp;
{$ELSE}
  THGrid(GetControl('FLISTE')).OnKeyUp := OnKeyUp;
{$ENDIF}
  BInsert := TToolBarButton97(GetControl('BINSERT',True));
  BDelete := TToolBarButton97(GetControl('BDELETE',True));
  BCalculer := TToolBarButton97(GetControl('BCALCULER',True));

  BInsert.OnClick := OnClickBInsert;
  BDelete.OnClick := OnClickBDelete;
  BCalculer.OnClick := OnClickBCalculer;

  MonPopUp := TPopUpMenu(GetControl('MONPOPUP',True));
  MonPopUp.Items[0].OnClick := OnClickBInsert;
  MonPopUp.Items[1].OnClick := OnClickBDelete;
  MonPopUp.Items[2].OnClick := OnClickBCalculer;
  AddMenuPop(MonPopUp,'','') ;

  PopUpImporter := TPopUpMenu(GetControl('POPUPIMPORTER',True));
  PopUpImporter.Items[0].OnClick := OnClickBRecuperation;
  PopUpImporter.Items[1].OnClick := OnClickBMiseAJour;

{$IFNDEF EAGLCLIENT}
  FListe := THDBGrid(GetControl('FLISTE'));
{$ELSE}
  FListe := THGrid(GetControl('FLISTE'));
{$ENDIF}
  FListe.OnDblClick := OnDblClickFListe;

  // On exclut le compte de capital de la MUL
  SetControlText('XX_WHERE','ICG_GENERAL <> "' + GetParamSocSecur('SO_ICCCOMPTECAPITAL','') + '"');

  // Société soumise à l' IS?
  Icc_Data.SoumisIs := TestImpositionSociete;

  FDateCtxDeb := IIF( CtxPCL in V_PGI.PgiContexte, VH^.CPExoRef.Deb, VH^.Encours.Deb );
  FDateCtxFin := IIF( CtxPCL in V_PGI.PgiContexte, VH^.CPExoRef.Fin, VH^.Encours.Fin );
  
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCGENERAUX.OnLoad;
begin
  inherited;
  CExoRefOuvert(True);
  SetFocusControl('ICG_GENERAL');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCGENERAUX.OnClickBCalculer(Sender: TObject);
var i: integer;
    S: string;
begin
  RechargeTheMulQ;

  if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
  begin
    S := '';
    if FListe.AllSelected then
    begin
      TheMulQ.First;
      while not TheMulQ.Eof do
      begin
        S := S + TheMulQ.FindField('ICG_GENERAL').AsString + ';';
        TheMulQ.Next;
      end;
    end
    else // Traitement <> si AllSelected car sinon ça merde, Si AllSelected alors NbSelected est faux, il vaut 0
    begin
      for i := 0 to FListe.nbSelected - 1 do
      begin
        FListe.GotoLeBookMark(i);
        S := S + GetField('ICG_GENERAL') + ';'
      end;
    end;

    if S <> '' then
      CPLanceFiche_ICCPARAMETRE(DateToStr(FDateCtxDeb) + ';' + DateToStr(FDateCtxFin) + ';' + S);
    FListe.ClearSelected;
  end
  else
    PGIINFO('Vous n''avez pas sélectionné de compte.', Ecran.Caption);
end;

{==============================================================================}
{========================= SUPRESSION DANS LE MUL =============================}
{==============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/09/2001
Modifié le ... : 10/05/2001
Description .. : Suppression dans le mullticritère de sélection des comptes ICC
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCGENERAUX.OnClickBDelete(Sender: TObject);
begin
  inherited;
  if not BlocageMonoPoste(True) then Exit;

  try
    if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
    begin
      if PgiAsk('Confirmez vous la suppression des comptes sélectionnés ?', Ecran.Caption) = MrYes then
      begin
        if Transactions(SuppressionCompte, 3) <> oeOk then
          MessageAlerte('Traitement annulé. Erreur lors de la suppression des comptes.');
      end;
    end
    else
      PGIINFO('Vous n''avez pas sélectionné de compte.', Ecran.Caption);
    FListe.ClearSelected;
  finally
     TFMUL(ECRAN).BChercheClick(nil);
    DeblocageMonoPoste(True);
  end;
end;

procedure TOF_ICCGENERAUX.SuppressionCompte;
var
  i: integer;
  LeCompte: string;
begin
  RechargeTheMulQ;

  if FListe.AllSelected then
  begin // On travaille sur tous les enregistrements de la Query
    TheMulQ.First;
    while not TheMulQ.Eof do
    begin
      LeCompte := TheMulQ.FindField('ICG_GENERAL').AsString;
      ExecuteSQL('DELETE FROM ICCGENERAUX WHERE ICG_GENERAL = "' + LeCompte + '"');
      ExecuteSQL('DELETE FROM ICCECRITURE WHERE ICE_GENERAL = "' + LeCompte + '"');
      TheMulQ.Next;
    end;
  end
  else
  begin // On travaille sur les enregistrements sélectionnés
    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookMark(i);
      LeCompte := GetField('ICG_GENERAL');
      ExecuteSQL('DELETE FROM ICCGENERAUX WHERE ICG_GENERAL = "' + LeCompte + '"');
      ExecuteSQL('DELETE FROM ICCECRITURE WHERE ICE_GENERAL = "' + LeCompte + '"');
    end;
  end;

  TheMulQ := nil;
end;

{==============================================================================}
{========================= IMPORTATION DES ECRITURES ==========================}
{==============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : Vérification avant appel de la procedure ImportationEcriture
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCGENERAUX.OnClickBRecuperation(Sender: TObject);
begin
  if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
  begin
    if PgiAsk('Confirmez-vous la récupération complète des écritures de la comptabilité ?', Ecran.Caption) = MrYes then
      ImportationEcriture(FALSE);
  end
  else
    PGIINFO('Vous n''avez pas sélectionné de compte.',Ecran.Caption);
  FListe.ClearSelected;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : Vérification avant appel de la procedure ImportationEcriture
Mots clefs ... :
*****************************************************************}

procedure TOF_ICCGENERAUX.OnClickBMiseAJour(Sender: TObject);
begin
  if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
  begin
    if PgiAsk('Confirmez vous la mise à jour avec les nouvelles écritures de la comptabilité ?', Ecran.Caption) = MrYes then
      ImportationEcriture(TRUE);
  end
  else
    PGIINFO('Vous n''avez pas sélectionné de compte.', Ecran.Caption);
  FListe.ClearSelected;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : Procedure d'imporation des écritures
Mots clefs ... :
*****************************************************************}

procedure TOF_ICCGENERAUX.ImportationEcriture(ModeMiseAJour: Boolean);
var j: integer;
begin
  if not BlocageMonoPoste(True) then Exit;
  try
    RechargeTheMulQ;
    BeginTrans;
    if FListe.AllSelected then
    begin
      { On travaille sur tous les enregistrements de la Query }
      TheMulQ.First;
      while not TheMulQ.Eof do
      begin
        if not TraitementEcriture(TheMulQ.FindField('ICG_GENERAL').AsString, FDateCtxDeb, FDateCtxFin, ModeMiseAJour) then
        begin
          RollBack;
          Exit;
        end;
        TheMulQ.Next;
      end;
    end
    else
    begin
      { On travaille sur les enregistrements sélectionnés }
      for j := 0 to FListe.nbSelected - 1 do
      begin
        FListe.GotoLeBookMark(j);
        if not TraitementEcriture(GetField('ICG_GENERAL'), FDateCtxDeb, FDateCtxFin, ModeMiseAJour) then
        begin
          RollBack;
          Exit;
        end;
      end;
    end;
    CommitTrans;

    if ModeMiseAJour then
      PgiInfo('Le traitement s''est correctement terminé. Aucune anomalie n''a été détectée.',Ecran.Caption)
    else
      PgiInfo('Le traitement s''est correctement terminé. Aucune anomalie n''a été détectée.',Ecran.Caption)

  finally
    DeblocageMonoPoste(True);
  end;
end;

procedure TOF_ICCGENERAUX.OnClose;
begin
  inherited;
  if Assigned(TheMulQ) then
    TheMulQ := nil;
end;

procedure TOF_ICCGENERAUX.OnClickBInsert(Sender: TObject);
begin
  CPLanceFiche_CreateICCGENERAUX;
  TFMUL(ECRAN).BChercheClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCGENERAUX.OnKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
  case Key of

    VK_F6: BCalculer.Click;

    VK_F11: TPopUpMenu(GetControl('MONPOPUP')).Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

    VK_DELETE: if Shift = [ssCtrl] then BDelete.Click;

    // GCO - 13/01/2005 - FQ 14953
    Ord('N') : if ssCtrl in Shift then  OnClickBInsert(nil);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ICCGENERAUX.OnDblClickFListe(Sender: TObject);
begin
  inherited;
  CPLanceFiche_FicheICCGENERAUX(GetField('ICG_GENERAL'));
  TFMUL(ECRAN).BChercheClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2006
Modifié le ... :   /  /
Description .. : FQ 18652
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCGENERAUX.RechargeTheMulQ;
begin
  if TheMulQ = nil then
  begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  end;
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  registerclasses([TOF_ICCGENERAUX]);
end.

