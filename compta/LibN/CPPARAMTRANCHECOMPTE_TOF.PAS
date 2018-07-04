{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPPARAMTRANCHECOMPTE ()
Mots clefs ... : TOF;CPPARAMTRANCHECOMPTE
*****************************************************************}
Unit CPPARAMTRANCHECOMPTE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     uTob,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     Vierge,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97 ;

Type
  TOF_CPPARAMTRANCHECOMPTE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      fAxe                : string;
      fGTranche           : THGrid;
      fGListe             : THGrid;
      fTLesRubriques      : TOB;
      fModifie            : boolean;
      fbChargementEnCours : boolean;
      procedure ChargeLesRubriques;
      procedure ChargeUneRubrique ( pRow : integer );
      function  EnregistreLaRubrique : boolean;
      function  ControleLaRubrique ( pstCode : string ) : boolean;
      procedure RetailleLesGrilles;
      procedure OnGrilleCellExit    (Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
      procedure OnListeRowEnter     (Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
      procedure OnGrilleSetEditText (Sender:TObject;ACol,ARow:Longint;const Value : widestring);
      procedure OnClickBInsert      (Sender : TObject);
      procedure OnClickBDelete      (Sender : TObject);
      procedure OnClickBEnregistrer (Sender : TObject);
      procedure OnClickBSupprimer   (Sender : TObject);
      procedure OnClickBInserer     (Sender : TObject);
      procedure OnChangeLibelle     (Sender : TObject);
      procedure OnFormKeyDown       (Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;

const
  cInclu = 1;
  cExclu = 2;

procedure CPLanceFiche_ParamTrancheCompte ( pstAxe : string );

Implementation

uses Windows;

procedure CPLanceFiche_ParamTrancheCompte ( pstAxe : string );
begin
  AGLLanceFiche ('CP','CPTRANCHECOMPTE','','',pstAxe);
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnArgument (S : String ) ;
begin
  Inherited ;

  fAxe := ReadTokenSt(S);
  fModifie := False;
  fbChargementEnCours := false;

  SetControlText ('RB_AXE', TraduireMemoire('Axe')+' '+fAxe);
  SetControlEnabled('RB_RUBRIQUE',False);

  fGTranche := THGrid (GetControl('GTRANCHE'));
  fGListe := THGrid (GetControl('GLISTE'));
  fGTranche.OnCellExit := OnGrilleCellExit;
  fGTranche.OnSetEditText := OnGrilleSetEditText;
  fGListe.OnRowEnter := OnListeRowEnter;

  TToolbarButton97(GetControl('BINSERT')).OnClick := OnClickBInsert;
  TToolbarButton97(GetControl('BDELETE')).OnClick := OnClickBDelete;
  TToolbarButton97(GetControl('BENREGISTRER')).OnClick := OnClickBEnregistrer;
  TToolbarButton97(GetControl('BSUPPRIMER')).OnClick := OnClickBSupprimer;
  TToolbarButton97(GetControl('BINSERER')).OnClick := OnClickBInserer;

  ECRAN.OnKeyDown := OnFormKeyDown;

  THcritMaskEdit(GetControl('RB_LIBELLE')).OnChange := OnChangeLibelle;

  fTLesRubriques := TOB.Create ('', nil, -1);

  SetControlText('RB_PREDEFINI','');

  ChargeLesRubriques;
  if fTLesRubriques.Detail.Count > 0 then ChargeUneRubrique (1)
  else OnClickBInsert(nil);
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnClose ;
begin
  fTLesRubriques.Free;
  Inherited ;
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMTRANCHECOMPTE.OnGrilleCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if ((ACol = cInclu) and (ARow = fGTranche.RowCount - 1) and (fGTranche.Cells[cInclu,ARow]<>'')) then
    fGTranche.RowCount := fGTranche.RowCount + 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 19/06/2007
Modifié le ... : 19/06/2007
Description .. : Chargement des rubriques "tranches de comptes" et mise à
Suite ........ : jour de la grille (liste des tranches, à droite de l'écran)
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPARAMTRANCHECOMPTE.ChargeLesRubriques;
var
  lstSQL : string;
begin
  fGListe.VidePile(True);
  fTLesRubriques.ClearDetail;
  lstSQL := 'SELECT * FROM RUBRIQUE WHERE RB_CLASSERUB="TRA" AND RB_AXE="'+fAxe+'"';
  lstSQL := lstSQL + ' AND (RB_PREDEFINI<>"DOS" OR (RB_PREDEFINI="DOS" AND RB_NODOSSIER="'+V_PGI.NoDossier+'"))';
  fTLesRubriques.LoadDetailDBFromSQL('RUBRIQUE',lstSQL);
  fTLesRubriques.PutGridDetail(fGListe,False,False,';RB_RUBRIQUE;RB_LIBELLE');
  RetailleLesGrilles;
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnListeRowEnter(Sender: TObject;
  Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if fModifie then
  begin
    case PGIAskCancel('Voulez-vous enregistrer les modifications ?') of
      mrYes : if not EnregistreLaRubrique then
        begin
          Cancel := True;
          exit;
        end;
      mrNo  : ;
      mrCancel :
        begin
          Cancel := True;
          exit;
        end;
    end;
  end;
  if fModifie then
  begin
    // Réinitialisation de l'interface
    fModifie := False;
    SetControlEnabled('RB_RUBRIQUE',False);
    SetControlEnabled('RB_LIBELLE',False);
    SetControlEnabled('BINSERT',True);
    SetControlEnabled('BDELETE',True);
    SetControlEnabled('BINSERER',True);
    SetControlEnabled('BSUPPRIMER',True);
    ChargeLesRubriques;
  end;
  ChargeUneRubrique ( Ou );
end;

procedure TOF_CPPARAMTRANCHECOMPTE.ChargeUneRubrique(pRow : integer);
var
  Sti, Sti1 : string;
  Ste, Ste1 : string;
begin
  fbChargementEnCours := True;
  fGTranche.VidePile(True);
  fGTranche.RowCount := 2;
  SetControlText ('RB_RUBRIQUE', fTLesRubriques.Detail[pRow-1].GetString('RB_RUBRIQUE'));
  SetControlText ('RB_LIBELLE', fTLesRubriques.Detail[pRow-1].GetString('RB_LIBELLE'));
  if fTLesRubriques.Detail[pRow-1].GetString('RB_PREDEFINI')='STD' then
    SetControlText('RB_PREDEFINI',TraduireMemoire('Standard'))
  else if fTLesRubriques.Detail[pRow-1].GetString('RB_PREDEFINI')='DOS' then
    SetControlText('RB_PREDEFINI',TraduireMemoire('Dossier'))
  else if fTLesRubriques.Detail[pRow-1].GetString('RB_PREDEFINI')='CEG' then
    SetControlText('RB_PREDEFINI',TraduireMemoire('Cegid'));

  Sti := fTLesRubriques.Detail[pRow-1].GetString('RB_COMPTE1');
  Ste := fTLesRubriques.Detail[pRow-1].GetString('RB_EXCLUSION1');
  Sti1 := ReadTokenSt(Sti);
  Ste1 := ReadTokenSt(Ste);
  while (Sti1 <> '') do
  begin
    fGTranche.Cells[cInclu,fGTranche.RowCount-1] := Sti1;
    fGTranche.Cells[cExclu,fGTranche.RowCount-1] := Ste1;
    Sti1 := ReadTokenSt (Sti);
    Ste1 := ReadTokenSt (Ste);
    fGTranche.RowCount := fGTranche.RowCount+1;
  end;
  if (fTLesRubriques.Detail[pRow-1].GetString('RB_PREDEFINI')='STD') and
    (not (ctxStandard in V_PGI.PGIContexte)) then
  begin
    SetControlEnabled ('RB_RUBRIQUE',False);
    SetControlEnabled ('RB_LIBELLE',False);
    SetControlEnabled ('GTRANCHE',False);
    SetControlEnabled ('BDELETE',False);
    SetControlEnabled ('BINSERER',False);
    SetControlEnabled ('BSUPPRIMER',False);
  end
  else
  begin
    SetControlEnabled ('RB_RUBRIQUE',False);
    SetControlEnabled ('RB_LIBELLE',True);
    SetControlEnabled ('GTRANCHE',True);
    SetControlEnabled ('BDELETE',True);
    SetControlEnabled ('BINSERER',True);
    SetControlEnabled ('BSUPPRIMER',True);
  end;
  fbChargementEnCours := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 19/06/2007
Modifié le ... :   /  /
Description .. : Suppression d'une rubrique
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPARAMTRANCHECOMPTE.OnClickBDelete(Sender: TObject);
var
  lstCode : string;
begin
  lstCode := GetControlText('RB_RUBRIQUE');

  if ( PGIAsk('Confirmez-vous la suppression de la rubrique '+lstCode+' ?') = mrYes) then
  begin
    // Vérification préalable : la rubrique ne doit pas être utilisée dans un paramétrage
    if ExisteSQL ('SELECT S_SECTION FROM SECTION WHERE S_TRANCHEGENEDE="'+lstCode+'"') then
    begin
      PGIBox('Cette rubrique est utilisée. Vous ne pouvez pas la supprimer.');
      exit;
    end;
    // Suppression de la rubrique
    if (ExecuteSQL ('DELETE FROM RUBRIQUE WHERE RB_RUBRIQUE="'+GetControlText('RB_RUBRIQUE')+'"')=1) then
      ChargeLesRubriques
    else PGIBox ('Erreur lors de la suppression de la rubrique.');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 19/06/2007
Modifié le ... :   /  /
Description .. : Création d'une rubrique
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPARAMTRANCHECOMPTE.OnClickBInsert(Sender: TObject);
begin
  if fModifie then
  begin
    case PGIAskCancel('Voulez-vous enregistrer les modifications ?') of
      mrYes : if not EnregistreLaRubrique then exit;
      mrNo  : ;
      mrCancel : exit;
    end;
  end;
  fModifie := True;
  SetControlEnabled('BINSERT',False);
  SetControlEnabled('BDELETE',False);
  SetControlEnabled('BINSERER',True);
  SetControlEnabled('BSUPPRIMER',True);
  SetControlEnabled('RB_RUBRIQUE',True);
  SetControlEnabled('RB_LIBELLE',True);
  SetControlText('RB_RUBRIQUE','');
  SetControlText('RB_LIBELLE','');
  SetControlEnabled ('GTRANCHE',True);
  fGTranche.VidePile(True);
  fGTranche.RowCount := 2;
  SetFocusControl('RB_RUBRIQUE');
  if ctxStandard in V_PGI.PGIContexte then SetControlText('RB_PREDEFINI',TraduireMemoire('Standard'))
  else SetControlText('RB_PREDEFINI',TraduireMemoire('Dossier'));
end;

function TOF_CPPARAMTRANCHECOMPTE.EnregistreLaRubrique : boolean;
var
  i : integer;
  lstCompte1, lstExclu1 : string;
  T : TOB;
  lstCode : string;
begin
  Result := False;
  lstCompte1 := '';
  lstExclu1 := '';
  lstCode := GetControlText('RB_RUBRIQUE');
  // Contrôle de la saisie
  if ControleLaRubrique ( lstCode ) then
  begin
    // Récupération des tranches de comptes
    for i := 1 to fGTranche.RowCount - 1 do
    begin
      if ((fGTranche.Cells[cInclu,i] <> '') or  (fGTranche.Cells[cExclu,i]<>'')) then
      begin
        lstCompte1 := lstCompte1+fGTranche.Cells[cInclu,i]+';';
        lstExclu1 := lstExclu1+fGTranche.Cells[cExclu,i]+';';
      end;
    end;
    // Initialisation de l'enregistrement rubrique
    T := TOB.Create ('RUBRIQUE',nil,-1);
    try
      T.PutValue('RB_RUBRIQUE',lstCode);
      T.PutValue('RB_LIBELLE',GetControlText('RB_LIBELLE'));
      T.PutValue('RB_COMPTE1',lstCompte1);
      T.PutValue('RB_EXCLUSION1',lstExclu1);
      T.PutValue('RB_DATEVALIDITE',iDate2099);
      T.PutValue('RB_TYPERUB','GEN');
      T.PutValue('RB_CLASSERUB','TRA');
      T.PutValue('RB_AXE',fAxe);
      if ctxStandard in V_PGI.PGIContexte then
      begin
        T.PutValue('RB_PREDEFINI','STD');
        T.PutValue('RB_NODOSSIER','000000');
      end
      else
      begin
        T.PutValue('RB_PREDEFINI','DOS');
        T.PutValue('RB_NODOSSIER',V_PGI.NoDossier);
      end;
      // Enregistrement de la rubrique
      T.SetAllModifie(True);
      T.InsertOrUpdateDB();
      Result := True;
    finally
      T.Free;
    end;
  end;
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnClickBEnregistrer(Sender: TObject);
begin
  EnregistreLaRubrique;
  // Réinitialisation de l'interface
  fModifie := False;
  SetControlEnabled('RB_RUBRIQUE',False);
  SetControlEnabled('RB_LIBELLE',False);
  SetControlEnabled('BINSERT',True);
  SetControlEnabled('BDELETE',True);
  ChargeLesRubriques;
  fGListe.Row := 1;
  ChargeUneRubrique(1);
end;

procedure TOF_CPPARAMTRANCHECOMPTE.RetailleLesGrilles;
begin
  fGTranche.ColWidths[0]:=20;
  fGTranche.ColWidths[1]:=150;
  fGTranche.ColWidths[2]:=150;

  fGListe.ColWidths[0]:=20;
  fGListe.ColWidths[1]:=80;
  fGListe.ColWidths[2]:=150;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 20/06/2007
Modifié le ... :   /  /
Description .. : Contrôle de la rubrique avant enregistrement
Mots clefs ... :
*****************************************************************}
function TOF_CPPARAMTRANCHECOMPTE.ControleLaRubrique (pstCode : string ) : boolean;
begin
  Result := False;
  // Est-ce que le code de la rubrique est correctement saisi ?
  if (Trim(pstCode)='') then
  begin
    PGIBox ('Code incorrect. Veuillez saisir un code pour identifier la rubrique.');
    exit;
  end;
  // Est-ce que le code est déjà utilisé dans une autre classe de rubrique ?
  if ExisteSQL ('SELECT 1 FROM RUBRIQUE WHERE RB_RUBRIQUE="'+pstCode+'" AND RB_CLASSERUB<>"TRA"') then
  begin
    PGIBox ('Ce code est déjà utilisé dans une autre classe de rubrique. Veuillez en choisir un autre.');
    exit;
  end;
  Result := True;
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnGrilleSetEditText (Sender:TObject;ACol,
  ARow:Longint;const Value : widestring);
begin
  fModifie := True;
  SetControlEnabled('RB_LIBELLE',True);
  SetControlEnabled('BINSERT',False);
  SetControlEnabled('BDELETE',False);
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnChangeLibelle(Sender: TObject);
begin
  if fbChargementEnCours then exit;
  fModifie := True;
  SetControlEnabled('BINSERT',False);
  SetControlEnabled('BDELETE',False);
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnClickBInserer(Sender: TObject);
begin
  if fGTranche.Row < 1 then exit;
  fGTranche.InsertRow(fGTranche.Row);
  fModifie := True;
  SetControlEnabled('BINSERT',False);
  SetControlEnabled('BDELETE',False);
  fGTranche.Row := fGTranche.Row - 1;   
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnClickBSupprimer(Sender: TObject);
begin
  if fGTranche.Row < 1 then exit;
  fGTranche.DeleteRow(fGTranche.Row);
  fModifie := True;
  SetControlEnabled('BINSERT',False);
  SetControlEnabled('BDELETE',False);
end;

procedure TOF_CPPARAMTRANCHECOMPTE.OnFormKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if ( csDestroying in Ecran.ComponentState ) then Exit ;
  if Key = VK_ESCAPE then fGTranche.CacheEdit;

  case Key of
    VK_F10 :
      begin
        if GetControlEnabled('BENREGISTRER') then
          OnClickBEnregistrer(nil);
        Key := 0;
      end;
    VK_INSERT :
      begin
        if GetControlEnabled('BINSERT') then
          OnClickBInsert(nil);
        Key := 0;
      end;
    VK_DELETE :
      begin
        if GetControlEnabled('BDELETE') then
          OnClickBDelete(nil);
        Key := 0;
      end;
    Ord('I')  :
      if Shift = [ssCtrl] then
      begin
        if GetControlEnabled('BINSERER') then
          OnClickBInserer (nil);
        Key := 0;
      end;
    Ord('S')  :
      if Shift = [ssCtrl] then
      begin
        if GetControlEnabled('BSUPPRIMER') then
          OnClickBSupprimer (nil);
        Key := 0;
      end;
  end;
end;

Initialization
  registerclasses ( [ TOF_CPPARAMTRANCHECOMPTE ] ) ;
end.
