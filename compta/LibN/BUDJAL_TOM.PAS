{ Unité : Source TOM de la TABLE : BUDJAL (BUDJAL)
--------------------------------------------------------------------------------------
    Version  |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               11/01/05  VL   Création de l'unité
06.50.001.015  24/08/05  JP   FQ 16003 : problème de dessins des grilles ne look 2003
06.50.001.016  25/08/05  JP   FQ 16514 : La gestion des périodes n'est pas prise en compte                
08.00.001.015  15/10/07  JP   FQ 21641 : Indice de liste hors limite, lorsque l'on paramètre les sections
                              FQ 21641 : Plantage lors de l'enregistrement sous Oracle (uniquement ?)
--------------------------------------------------------------------------------------}

Unit BUDJAL_TOM ;

Interface

Uses Windows,
     {$IFDEF VER150} Variants, {$ENDIF}
     StdCtrls, Controls, Classes, Graphics, Buttons,
{$IFDEF EAGLCLIENT}
     eFiche, eFichList, Maineagl,
{$ELSE}
     db, HDB, HREGCPTE, 
     {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF},
     Fiche, FichList, FE_Main,
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
     PLANBUDJAL_TOF, // PlanBudJal
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HTB97, UTOM, utilPGI, AglInit, Grids,
     Hcompte, LookUp, Choix, Menus, BudJalSP, HStatus,Ent1,
     UTob ;

procedure FicheBudjal(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
procedure FicheBudjalMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);

Type
  TOM_BUDJAL = Class (TOM)
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
    Fiche : TFFiche;
    FListe, FListeG, FListeS : THGrid;
{$IFDEF EAGLCLIENT}
    BJ_CATEGORIE : THValComboBox;
    BJ_SOUSPLAN : THEdit;
    BJ_EXODEB : THValComboBox;
    BJ_EXOFIN : THValComboBox;
    {JP 25/08/05 : FQ 16514}
    BJ_COMPTEURNORMAL : THValComboBox;
    BJ_NATJAL : THValComboBox;
{$ELSE}
    BJ_CATEGORIE : THDBValComboBox;
    BJ_SOUSPLAN : THDBEdit;
    BJ_EXODEB : THDBValComboBox;
    BJ_EXOFIN : THDBValComboBox;
    {JP 25/08/05 : FQ 16514}
    BJ_COMPTEURNORMAL : THDBValComboBox;
    BJ_NATJAL : THDBValComboBox;
{$ENDIF}
    {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin
     JP 26/12/07 : Remplacement des DBCombo par de simples combo en 2/3}
    BJPERDEB : THValComboBox;
    BJPERFIN : THValComboBox;

    BUp, BDown, BAjout, BDel, BCOPIECOMPLEMENT : TToolBarButton97;

    SectionOk : Boolean;
    MsgBox : THMsgBox;
    Mode : TActionFiche;
    LaPage : Integer;
    Lequel,LeCompte : String;
    LesModif : String ;
    Modifier,BGModifier,BSModifier : Boolean ;
    MemoCod : String ;
    IsMouvementer : Boolean ;
    Lefb : TFichierBase ;
    FAvertir : Boolean ;
    procedure BImprimerClick(Sender: TObject);
    procedure BAjoutClick(Sender: TObject);
    procedure BDelClick(Sender: TObject);
    procedure BUpClick(Sender: TObject);
    procedure BDownClick(Sender: TObject);
    procedure BCopierJalClick(Sender: TObject);
    procedure BCopieCptClick(Sender: TObject);
    procedure BCopieSectClick(Sender: TObject);
    procedure BCopieDeuxClick(Sender: TObject);

    procedure BFerClick(Sender: TObject);
    procedure BValClick(Sender: TObject);

    procedure GbCpteEnter(Sender: TObject);

    procedure FListeGSorted(Sender: TObject);
    procedure FListeSSorted(Sender: TObject);

    procedure BZoomSousPlanClick(Sender: TObject);
    procedure BJ_SOUSPLANKeyPress(Sender: TObject; var Key: Char);

    procedure BJ_EXODEBChange(Sender: TObject);
    procedure BJ_EXOFINChange(Sender: TObject);
    procedure BJ_AXEChange(Sender: TObject);
    procedure BJ_AXEClick(Sender: TObject);
    procedure BJ_PERDEBChange(Sender: TObject);
    procedure BJ_PERFINChange(Sender: TObject);
    procedure BJ_NATJALChange(Sender: TObject);
    procedure BJ_CATEGORIEChange(Sender: TObject);

    function  CodeValide : Boolean;
    procedure AffecteLe(Champ, Valeur : string);
    procedure InitModifEnSerie(StModif : string);
    procedure LitBudgensBudsects(Stc : String ; G : THGrid ; PremFois : Boolean);
    function  EcritBudgensBudsects(G : THGrid ; Var Ind : Integer) : String;
    function  CpteAttenteOk ( FicBase : TFichierBase) : Boolean;
    procedure GriseDegriseControl;
    function  DoublonExiste(St : String ; G : THGrid) : Boolean;
    function  ExoOuv : Boolean;
    procedure PositionnePerDeb;
    procedure PositionnePerFin;
    function  TrouveCpteMvt(G : THGrid) : Boolean;
    function  CoherenceSection : Boolean;
    procedure GenereCroisement(CodeACreer,CodeChoisi : String);
    procedure DesactiveLesControles;
    procedure CopieCatSurJournaux;
    procedure BCopieListe(Qui : Integer);
    procedure LibelleCatBud(CodeCat,SousPlanBud : String ; Var LibBJ,AbrBJ : String);
    procedure UpdateDate;
    procedure ListeWidth;
  public
    {JP 24/08/05 : FQ 16003 : lors de l'affichage de la fiche les colonnes des grilles
               sont redimensionnées sur la base de la taille de la fiche non sur
               celle de la grille}
    OnAncetreEcranResize : TNotifyEvent;
    procedure OnEcranResize(Sender : TObject);
    {FIN JP 24/08/05 : FQ 16003}
  end ;

function EstMouvementeBudjal(Lejal : String) : Boolean;
function ExisteBudJal(const Lejal : String) : Boolean;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  Paramsoc;


procedure FicheBudjal(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
begin
  if ((Comment in [taCreat..taCreatOne])=False) then if not ExisteBudJal(Compte) then exit;
  if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit;
  AGLLanceFiche('CP', 'CPBUDJAL', '', Compte, ActionToString(Comment)+';'+Axe+';'+Compte+';'+IntToStr(QuellePage));
end;

procedure FicheBudjalMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
begin
  if ((Comment in [taCreat..taCreatOne])=False) then if not ExisteBudJal(Lequel) then exit;
  if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit;
  AGLLanceFiche('CP', 'CPBUDJAL', '', Lequel, TAToStr(Comment)+';'+Axe+';'+Lequel+';'+IntToStr(QuellePage)+';'+LesModif);
end;

function EstMouvementeBudjal(Lejal : String) : Boolean ;
begin
  Result := ExisteSQL('SELECT BJ_BUDJAL FROM BUDJAL WHERE BJ_BUDJAL="'+Lejal+'" AND '+
              '(EXISTS(SELECT BE_BUDJAL FROM BUDECR WHERE BE_BUDJAL="'+Lejal+'"))');
end;

function ExisteBudJal(const Lejal : String) : Boolean;
begin
  Result := True;
  if Lejal= '' then exit;
  if not ExisteSQL('SELECT BJ_BUDJAL FROM BUDJAL WHERE BJ_BUDJAL="'+Lejal+'"') then begin
    MessageAlerte('L''enregistrement est inaccessible.');
    Result := False;
  end;
end;

procedure TOM_BUDJAL.OnNewRecord ;
begin
  Inherited ;

  UpdateDate;
  IsMouvementer := False;
  MemoCod := '';
  GriseDegriseControl;
  FListeG.VidePile(False);
  FListeS.VidePile(False);
  SetField('BJ_AXE', 'A1');
  BJ_AXEChange(Nil) ;
  SetField('BJ_NATJAL', 'CHA');
  BJ_NATJALChange(Nil);
  BJ_EXODEB.Value := VH^.Encours.Code;
  BJ_EXOFIN.Value := VH^.Encours.Code;
  SetField('BJ_EXODEB', VH^.Encours.Code);
  SetField('BJ_EXOFIN', VH^.Encours.Code);
  PositionnePerDeb;
  PositionnePerFin;
  if BJ_COMPTEURNORMAL.Values.Count>0 then BJ_COMPTEURNORMAL.Value := BJ_COMPTEURNORMAL.Values[0];
  SetControlEnabled('BJ_BUDJAL', True);
  Fiche.Pages.ActivePage := Fiche.Pages.Pages[0];
  SetFocusControl('BJ_BUDJAL');
  SetControlEnabled('BCOPIERJAL', True);
end ;

procedure TOM_BUDJAL.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BUDJAL.OnUpdateRecord ;
var
  Ind : Integer;
begin
  Inherited ;

  if Not CodeValide then begin
    LastError :=1;
    exit;
  end;

  if DS.state in [dsEdit,dsInsert] then begin
    if GetControlText('BJ_COMPTEURNORMAL') = '' then begin
      Msgbox.Execute(25,'',''); // Vous devez renseigner le compteur.
      LastError :=1;
      Exit;
    end;

    if Not CpteAttenteOk(fbBudGen) or Not CpteAttenteOk(Lefb) or Not CoherenceSection then begin
      LastError :=1;
      exit;
    end;

    if Not ExoOuv then begin
      Fiche.Pages.ActivePage := Fiche.Pages.Pages[0];
      SetFocusControl('BJ_EXODEB');
      Msgbox.Execute(12,'',''); // Cet exercice n''est pas ouvert.
      LastError :=1;
      Exit ;
    end;

    {JP 15/10/07 : FQ 21641 : Déplacer depuis le OnAfterUpdate}
    if BGModifier or BSModifier then CopieCatSurJournaux;
    Ind := 1;
    SetField('BJ_BUDGENES', EcritBudgensBudsects(FListeG,Ind));
    If Ind > 0 then SetField('BJ_BUDGENES2', EcritBudgensBudsects(FListeG,Ind));
    Ind := 1;
    SetField('BJ_BUDSECTS', EcritBudgensBudsects(FListeS,Ind));
    If Ind > 0 Then SetField('BJ_BUDSECTS2', EcritBudgensBudsects(FListeS,Ind));

    // 13343
    if GetField('BJ_BUDGENES' ) = '' then SetField('BJ_BUDGENES' , ' ');
    if GetField('BJ_BUDGENES2') = '' then SetField('BJ_BUDGENES2', ' ');
    if GetField('BJ_BUDSECTS' ) = '' then SetField('BJ_BUDSECTS' , ' ');
    if GetField('BJ_BUDSECTS2') = '' then SetField('BJ_BUDSECTS2', ' ');
    {FIN : JP 15/10/07 : FQ 21641 : Déplacer depuis le OnAfterUpdate}

    FAvertir := True;

  end;
end ;

procedure TOM_BUDJAL.OnAfterUpdateRecord ;
begin
  Inherited ;

end ;

procedure TOM_BUDJAL.OnAfterDeleteRecord ;
begin
  Inherited ;

  FAvertir := True;
end ;

procedure TOM_BUDJAL.OnLoadRecord ;
var
  Lib,Abr : String;
begin
  Inherited ;

  UpdateDate;

  FListeS.SortedCol := -1;
  FListeG.SortedCol := -1;

  SetControlEnabled('BCOPIERJAL', DS.State in [dsInsert]);

  if LeCompte = GetField('BJ_BUDJAL') then begin
    if Mode=taConsult then DesactiveLesControles;
    Exit ;
  end;

  if (GetField('BJ_LIBELLE') = '') And (GetField('BJ_SOUSPLAN')<>'') then begin
    LibelleCatBud(GetField('BJ_CATEGORIE'), GetField('BJ_SOUSPLAN'), Lib, Abr);
    SetField('BJ_LIBELLE', Lib);
    SetField('BJ_ABREGE', Abr);
  end;
  LeCompte := GetField('BJ_BUDJAL');

  if DS.State<>dsInsert then begin
    if MemoCod<>GetField('BJ_BUDJAL') then begin
      IsMouvementer := EstMouvementeBudjal(GetField('BJ_BUDJAL'));
      LitBudgensBudsects(Trim(GetField('BJ_BUDGENES')), FListeG, True);
      LitBudgensBudsects(Trim(GetField('BJ_BUDGENES2')), FListeG, False);
      LitBudgensBudsects(Trim(GetField('BJ_BUDSECTS')), FListeS, True);
      LitBudgensBudsects(Trim(GetField('BJ_BUDSECTS2')), FListeS, False);
      {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
      ListePeriode(BJ_EXODEB.Value, BJPERDEB.Items, BJPERDEB.Values, True);
      ListePeriode(BJ_EXOFIN.Value, BJPERFIN.Items, BJPERFIN.Values, False);
      BJPERDEB.ItemIndex := BJPERDEB.Values.IndexOf(DateToStr(GetField('BJ_PERDEB')));
      BJPERFIN.ItemIndex := BJPERFIN.Values.IndexOf(DateToStr(GetField('BJ_PERFIN')));

      MemoCod := GetField('BJ_BUDJAL');
    end;
  end;
  BJ_AXEChange(Nil);
  GriseDegriseControl;

  case Mode Of
    taConsult : DesactiveLesControles;
    taModifEnSerie : InitModifEnSerie(LesModif);
  end;
end ;

procedure TOM_BUDJAL.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

// 1° : TActionFiche
// 2° : Axe
// 3° : Lequel ou Compte
// 4° : QuellePage
// 5° : LesModif
procedure TOM_BUDJAL.OnArgument ( S: String ) ;
begin
  Inherited ;

  {JP 24/08/05 : FQ 16003 : lors de l'affichage de la fiche les colonnes des grilles
                 sont redimensionnées sur la base de la taille de la fiche non sur
                 celle de la grille}
  if Assigned(TFFiche(Ecran).OnResize) then OnAncetreEcranResize := TFFiche(Ecran).OnResize;
  TFFiche(Ecran).OnResize := OnEcranResize;

  // Paramètres
  Mode := StrToTA(ReadTokenST(S));
  ReadTokenST(S); // Axe
  Lequel := ReadTokenST(S);
  LaPage := READTOKENI(S);
  LesModif := ReadTokenST(S);

  MsgBox := THMsgBox.Create(Ecran);
  MsgBox.Mess.Add('0;Budgets;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;C;');
  MsgBox.Mess.Add('1;Budgets;Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;');
  MsgBox.Mess.Add('2;Budgets;Vous devez renseigner le code.;W;O;O;O;');
  MsgBox.Mess.Add('3;Budgets;Vous devez renseigner le libellé.;W;O;O;O;');
  MsgBox.Mess.Add('4;Budgets;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
  MsgBox.Mess.Add('L''enregistrement est inaccessible.');
  MsgBox.Mess.Add('Code');
  MsgBox.Mess.Add('Libellé');
  MsgBox.Mess.Add('Axe');
  MsgBox.Mess.Add('9;Budgets;Confirmez-vous la suppression de cette valeur ?;Q;YN;N;N;');
  MsgBox.Mess.Add('10;Budgets;Vous devez renseigner le compte d''attente.;W;O;O;O;');
  MsgBox.Mess.Add('11;Budgets;Le compte d''attente renseigné n''existe pas.;W;O;O;O;');
  MsgBox.Mess.Add('12;Budgets;Cet exercice n''est pas ouvert.;E;O;O;O;');
  MsgBox.Mess.Add('13;Budgets;Vous ne pouvez pas supprimer cette valeur. Des écritures lui sont rattachées.;W;O;O;O;');
  MsgBox.Mess.Add('14;Budgets;Vous devez renseigner la section d''attente.;W;O;O;O;');
  MsgBox.Mess.Add('15;Budgets;La section d''attente renseignée n''existe pas.;W;O;O;O;');
  MsgBox.Mess.Add('16;Budgets;Des sections n''appartiennent pas à l''axe choisi.;E;O;O;O;');
  MsgBox.Mess.Add('Choix d''un budget à recopier');
  MsgBox.Mess.Add('18;Budgets;Confirmez-vous la recopie des comptes du budget sélectionné ?;Q;YN;N;N;');
  MsgBox.Mess.Add('19;Budgets;Confirmez-vous la recopie du budget sélectionné ?;Q;YN;N;N;');
  MsgBox.Mess.Add('20;Budgets;Opération impossible : Vous n''avez pas renseigné le code identifiant du budget.;W;O;O;O;');
  MsgBox.Mess.Add('Génération des croisements budgétaires en cours...');
  MsgBox.Mess.Add('22;Budgets;Désirez-vous apporter les modifications des comptes / sections à d''autres budgets de la catégorie concernée ?;Q;YN;N;N;');
  MsgBox.Mess.Add('Choix des journaux budgétaires à modifier');
  MsgBox.Mess.Add('Saisie du budget en');
  MsgBox.Mess.Add('25;Budgets;Vous devez renseigner le compteur.;W;O;O;O;');

  Fiche := TFFiche(Ecran);
  FListe := THGrid(GetControl('FLISTE', True));
  FListeG := THGrid(GetControl('FLISTEG', True));
  FListeS := THGrid(GetControl('FLISTES', True));

{$IFDEF EAGLCLIENT}
  BJ_CATEGORIE := THValComboBox(GetControl('BJ_CATEGORIE', True));
  BJ_SOUSPLAN := THEdit(GetControl('BJ_SOUSPLAN', True));
  BJ_EXODEB := THValComboBox(GetControl('BJ_EXODEB', True));
  BJ_EXOFIN := THValComboBox(GetControl('BJ_EXOFIN', True));
  {JP 25/08/05 : FQ 16514}
  BJ_COMPTEURNORMAL := THValComboBox(GetControl('BJ_COMPTEURNORMAL', True));
  BJ_NATJAL := THValComboBox(GetControl('BJ_NATJAL', True));
{$ELSE}
  BJ_CATEGORIE := THDBValComboBox(GetControl('BJ_CATEGORIE', True));
  BJ_SOUSPLAN := THDBEdit(GetControl('BJ_SOUSPLAN', True));
  BJ_EXODEB := THDBValComboBox(GetControl('BJ_EXODEB', True));
  BJ_EXOFIN := THDBValComboBox(GetControl('BJ_EXOFIN', True));
  {JP 25/08/05 : FQ 16514}
  BJ_COMPTEURNORMAL := THDBValComboBox(GetControl('BJ_COMPTEURNORMAL', True));
  BJ_NATJAL := THDBValComboBox(GetControl('BJ_NATJAL', True));
{$ENDIF}
  {JP 26/12/07 : Remplacement des DBCombo par de simple combo.}
  BJPERDEB  := THValComboBox(GetControl('BJPERDEB', True));
  BJPERFIN  := THValComboBox(GetControl('BJPERFIN', True));

  Fiche.BImprimer.OnClick := BImprimerClick;
  Fiche.BImprimer.Visible := True;

  FListeG.OnSorted := FListeGSorted;
  FListeS.OnSorted := FListeSSorted;

  BJ_CATEGORIE.OnChange := BJ_CATEGORIEChange;
  BJ_SOUSPLAN.OnElipsisClick := BZoomSousPlanClick;
  BJ_SOUSPLAN.OnKeyPress := BJ_SOUSPLANKeyPress;
  BJ_EXODEB.OnChange := BJ_EXODEBChange;
  BJ_EXOFIN.OnChange := BJ_EXOFINChange;
  TComboBox(GetControl('BJ_AXE', True)).OnChange := BJ_AXEChange;
  TComboBox(GetControl('BJ_AXE', True)).OnClick := BJ_AXEClick;
  {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
  BJPERDEB.OnChange := BJ_PERDEBChange;
  BJPERFIN.OnChange := BJ_PERFINChange;
  BJ_NATJAL.OnChange := BJ_NATJALChange;

  BUp := TToolBarButton97(GetControl('BUP', True));
  BDown := TToolBarButton97(GetControl('BDOWN', True));
  BAjout := TToolBarButton97(GetControl('BAJOUT', True));
  BDel := TToolBarButton97(GetControl('BDEL', True));
  BCOPIECOMPLEMENT := TToolBarButton97(GetControl('BCOPIECOMPLEMENT', True));

  BUp.OnClick := BUpClick;
  BDown.OnClick := BDownClick;
  BAjout.OnClick := BAjoutClick;
  BDel.OnClick := BDelClick;

  TToolBarButton97(GetControl('BFER', True)).OnClick := BFerClick;
  TToolBarButton97(GetControl('BVAL', True)).OnClick := BValClick;
  TGroupBox(GetControl('GBCPTE', True)).OnEnter := GbCpteEnter;
  TToolBarButton97(GetControl('BCOPIERJAL', True)).OnClick := BCopierJalClick;

  TMenuItem(GetControl('BCOPIECPT', True)).OnClick := BCopieCptClick;
  TMenuItem(GetControl('BCOPIESECT', True)).OnClick := BCopieSectClick;
  TMenuItem(GetControl('BCOPIEDEUX', True)).OnClick := BCopieDeuxClick;

  // FormShow
  LeCompte:=W_W;
  SetControlVisible('GBTAB', False);
  MemoCod := '';
  FListe.MultiSelect := True;

  Fiche.Pages.ActivePageIndex := LaPage;
  case Mode Of
    taConsult : DesactiveLesControles;
    taCreat..taCreatOne : begin
      SetField('BJ_BUDJAL', Lequel);
      Fiche.bDefaire.Enabled:=False;
    end;
  end;
  Ecran.HelpContext := 15151200;
end ;

procedure TOM_BUDJAL.OnClose ;
begin
  Inherited ;
  if FAvertir then begin
    AvertirTable('ttBudJal');
    AvertirTable('ttBudJalSais');
    AvertirTable('TTBUDJALSANSCAT');
    AvertirTable('TTCATJALBUD');
  end;
  MsgBox.Free;
end ;

procedure TOM_BUDJAL.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BUDJAL.LibelleCatBud(CodeCat, SousPlanBud: String; var LibBJ, AbrBJ: String);
var
  k,l : Integer ;
  St,StSousPlan,St2,Code,Lib : String ;
  SousPlan : TSousPlan ;
  SPC : TSousPlanCat ;
  Cat : TUneCatBud ;
begin
  LibBJ := '';
  AbrBJ := '';
  SPC := SousPlanCat(CodeCat, True);
  Cat := QuelleCatBud(CodeCat);
  if Cat.Code='' then Exit;

  for l:=1 To MaxSousPlan do begin
    St := Cat.SurJal[l];
    SousPlan := SPC[l];
    if St<>'' then begin
      StSousPlan := ReadTokenSt(SousPlanBud);
      if StSousPlan<>'' then begin
        for k:=0 To SousPlan.ListeSP.Count-1 do begin
          St2 := SousPlan.ListeSP.Strings[k];
          Code := ReadTokenSt(St2);
          Lib := ReadTokenSt(St2);
          if Code=StSousPlan then begin
            LibBJ := LibBJ+Lib;
            AbrBJ := AbrBJ+Code;
            Break;
          end;
        end;
      end;
    end;
  end;
end;

procedure TOM_BUDJAL.DesactiveLesControles;
begin
  // Onglet compléments
  SetControlEnabled('GBCPTE', True);
  FListeG.Enabled := True;
  FListeS.Enabled := True;

  SetControlEnabled('PANEL1', False);
  SetControlEnabled('BUP', False);
  SetControlEnabled('BDOWN', False);
  SetControlEnabled('BAJOUT', False);
  SetControlEnabled('BDEL', False);
  SetControlEnabled('BCCOPIECOMPLEMENT', False);
  FListeG.SortEnabled := False;
  FListeS.SortEnabled := False;
end;

procedure TOM_BUDJAL.BImprimerClick(Sender: TObject);
begin
{$IFNDEF IMP}
PlanBudJal(GetField('BJ_AXE'), GetField('BJ_BUDJAL'), True);
{$ENDIF}
end;

function TOM_BUDJAL.CodeValide: Boolean;
begin
  Result := False;
  if DS.state in [dsInsert] then begin
    if Presence('BUDJAL','BJ_BUDJAL', GetField('BJ_BUDJAL')) then begin
      Fiche.Pages.ActivePage := Fiche.Pages.Pages[0];
      SetFocusControl('BJ_BUDJAL');
      MsgBox.Execute(4,'',''); // Le code que vous avez saisi existe déjà. Vous devez le modifier.
      Exit;
    end;
  end;
  Result := True;
end;

procedure TOM_BUDJAL.InitModifEnSerie(StModif: string);
var
  St,Champ, Valeur : string;
  i : integer;
  B : TBitBtn;
  PBouton : TToolWindow97;
begin
  if DS.State=dsBrowse then DS.Edit;

  While StModif<>'' do begin
    St := ReadTokenSt(StModif);
    i:=Pos('=',St); if i>0 then Champ := Trim(Copy(St,1,i-1));
    i:=Pos('"',St); if i>0 then St := Trim(Copy(St,i+1,Length(St)));
    i:=Pos('"',St); if i>0 then Valeur := Trim(Copy(St,1,i-1));
    AffecteLe(Champ,Valeur);
  end;

  PBouton := TToolWindow97(GetControl('PBOUTON', True));
  for i:=0 to PBouton.ControlCount-1 do begin
    if PBouton.Controls[i] is TBitBtn then begin
      B := TBitBtn(PBouton.Controls[i]);
      if ((UpperCase(B.Name)<>'BVALIDER') and
          (UpperCase(B.Name)<>'BFERME') and
          (UpperCase(B.Name)<>'BHELPBTN')) then B.Enabled := False;
    end;
  end;
end;

procedure TOM_BUDJAL.AffecteLe(Champ, Valeur: string);
var
  C : TControl;
begin
  C := GetControl(Champ);

  {$IFDEF EAGLCLIENT}
  if (C is TCheckBox) or (C is THValComboBox) or (C is TEdit) or (C is THEdit) then
  {$ELSE}
  if (C is THDBCheckBox) or (C is THDBValComboBox) or (C is THDBEdit) or (C is THDBCpteEdit) then
  {$ENDIF}
    begin
    SetField(Champ, Valeur);
    TEdit(C).Font.Color := clRed;
    end
  {$IFDEF EAGLCLIENT}
  else if (C is THSpinEdit) then
  {$ELSE}
  else if (C is THDBSpinEdit) then
  {$ENDIF}
    begin
    SetField(Champ,Valeur);
    THSpinEdit(C).Font.Color := clRed;
  end;
end;

procedure TOM_BUDJAL.BAjoutClick(Sender: TObject);
var
  T : Tob;
begin
  T := TOB.Create('_BUDGENE', nil, -1);
  try
    if Ecran.ActiveControl = FListeG then begin
      T.LoadDetailFromSQL('SELECT BG_BUDGENE, BG_LIBELLE FROM BUDGENE WHERE BG_ATTENTE="-" ORDER BY BG_BUDGENE');
      T.PutGridDetail(FListe, True, True, 'BG_BUDGENE;BG_LIBELLE', True);
      FListe.ColWidths[0] := 90;
      FListe.ColWidths[1] := 275;
      FListe.ColCount := 2;
      BGModifier:=True ;
      SectionOk := False;
    end
    else if Ecran.ActiveControl = FListeS then begin
      FListe.ColCount := 3; {JP 25/08/05 : FQ 16514}
      T.LoadDetailFromSQL('SELECT BS_BUDSECT,BS_LIBELLE,BS_AXE FROM BUDSECT WHERE BS_AXE="'+GetField('BJ_AXE')+'" AND BS_ATTENTE="-" ORDER BY BS_BUDSECT');
      T.PutGridDetail(FListe, True, True, 'BS_BUDSECT;BS_LIBELLE;BS_AXE', True);
      FListe.ColWidths[0] := 90;
      FListe.ColWidths[1] := 240;
      FListe.ColWidths[2] := 35;
      FListe.Cells[2, 0] := MsgBox.Mess[8]; // Axe
      BSModifier:=True ;
      SectionOk := True;
    end
    else
      Exit;
  finally
    T.Free;
  end;

  FListe.Cells[0, 0] := MsgBox.Mess[6]; // Code
  FListe.Cells[1, 0] := MsgBox.Mess[7]; // Libellé
  SetControlVisible('GBCPTE', False);
  SetControlVisible('GBTAB', True);
  SetControlEnabled('PBOUTON', False);
  Fiche.Pages.Pages[0].TabVisible := False;
  Fiche.Pages.Pages[2].TabVisible := False;

  if not(DS.State in [dsInsert, dsEdit]) then DS.Edit;
end;

procedure TOM_BUDJAL.BFerClick(Sender: TObject);
begin
  SetControlVisible('GBTAB', False);
  SetControlVisible('GBCPTE', True);
  SetControlEnabled('PBOUTON', True);
  Fiche.Pages.Pages[0].TabVisible := True;
  Fiche.Pages.Pages[2].TabVisible := True;
  Fiche.Pages.ActivePageIndex := 1;
end;

procedure TOM_BUDJAL.BValClick(Sender: TObject);
var
  i,NbLig : Integer;
  St,St1 : String;
  G : THGrid;
begin
  NbLig := Fliste.NbSelected;
  if NbLig<=0 then begin
    BFerClick(Nil);
    Exit;
  end;

  St := '';
  Modifier := False;

  for i:=0 to NbLig-1 do begin
    Fliste.GotoLeBookmark(i);
    if SectionOk then St := St + FListe.Cells[2, FListe.Row] + FListe.Cells[0, FListe.Row] + ';'
                 else St := St + FListe.Cells[0, FListe.Row] + ';' ;
  end;
  
  if SectionOk then G:=FListeS
               else G:=FListeG;

  while St<>'' do begin
    St1 := ReadTokenSt(St);
    if G=FListeS then
       if Copy(St1,1,2)<>GetField('BJ_AXE') then Continue;
    if DoublonExiste(St1,G) then Continue;
    if G.Cells[0,G.RowCount-1]<>'' then G.RowCount := G.RowCount+1;
    if G=FListeS then begin
      G.Cells[1,G.RowCount-1] := Copy(St1,1,2);
      Delete(St1,1,2);
    end;
    G.Cells[0,G.RowCount-1] := St1;
    Modifier := True;
  end;
  BFerClick(Nil);

  {JP 15/10/07 : FQ 21641 : Pour éviter un indice de liste hors limite quand l'on passe des sections
                 aux généraux : auparavant, la grille n'était pas correctement réinitialisée}
  for i := 0 to Fliste.RowCount - 1 do
    if Fliste.IsSelected(i) then Fliste.FlipSelection(i);
  Fliste.Row := 2;
end;

procedure TOM_BUDJAL.LitBudgensBudsects(Stc: String; G: THGrid; PremFois: Boolean);
begin
  if PremFois then G.VidePile(False)
              else if Stc='' then Exit;

  while Stc<>'' do begin
    if G.Cells[0,G.RowCount-1]<>'' then G.RowCount := G.RowCount+1;
    if UpperCase(G.Name)='FLISTES' then G.Cells[1,G.RowCount-1] := GetField('BJ_AXE');
    G.Cells[0,G.RowCount-1] := ReadTokenSt(Stc);
  end;


end;

function TOM_BUDJAL.EcritBudgensBudsects(G: THGrid; var Ind: Integer): String;
var
  St,St1 : String;
  i,j : Integer;
begin
  St := '';
  j := 0;
  if Ind>0 then begin
    for i:=Ind to G.RowCount-1 do begin
      if (i=Ind) And (G.Cells[0,i]='') then Break;
      St1 := St + G.Cells[0,i] + ';';
      if Length(St1)<=2000 then St := St1 else begin
        j := i;
        Break;
      end;
    end;
  end;

  Result := St;
  Ind := j;
end;

function TOM_BUDJAL.TrouveCpteMvt(G: THGrid): Boolean;
var
  QLoc : TQuery;
  Champ,St : String;
begin
  Result := False;
  if Not IsMouvementer then Exit;
  if UpperCase(G.Name)='FLISTEG' then Champ := 'BE_BUDGENE'
                                 else Champ := 'BE_BUDSECT';
  QLoc := OpenSQL('SELECT '+Champ+' FROM BUDECR WHERE BE_BUDJAL="'+GetField('BJ_BUDJAL')+'" AND '+Champ+'="'+G.Cells[0,G.Row]+'"',True) ;
  St := QLoc.Fields[0].AsString;
  Ferme(QLoc);
  SourisNormale;
  if St='' then Exit ;
  Result := True;
end;

procedure TOM_BUDJAL.BDelClick(Sender: TObject);
var
  G : THGrid;
  Cpte : String;
begin

  if Ecran.ActiveControl = FListeG then begin
    BGModifier := True;
    G := FListeG;
    Cpte := 'CX_COMPTE';
  end
  else if Ecran.ActiveControl = FListeS then begin
    BSModifier := True;
    G := FListeS;
    Cpte := 'CX_SECTION';
  end
  else
    Exit;

  if (G.Row=1) And (G.Cells[0,G.Row]='') then begin
    BGModifier := False;
    BSModifier := False;
    Exit;
  end;
  if MsgBox.Execute(9,'','')<>mrYes then begin // Confirmez-vous la suppression de cette valeur ?
    BGModifier := False;
    BSModifier := False;
    Exit;
  end;

  if TrouveCpteMvt(G) Then begin
    MsgBox.Execute(13,'',''); // Vous ne pouvez pas supprimer cette valeur. Des écritures lui sont rattachées.
    Exit;
  end;

  Modifier := True;
  ExecuteSql('DELETE FROM CROISCPT WHERE CX_JAL="'+GetField('BJ_BUDJAL')+'" AND CX_TYPE="BUD" AND '+Cpte+'="'+G.Cells[0,G.Row]+'"');
  if GetField('BJ_CATEGORIE')<>'' Then
    ExecuteSql('DELETE FROM CROISCPT WHERE CX_JAL="'+GetField('BJ_CATEGORIE')+'" AND CX_TYPE="BUD" AND '+Cpte+'="'+G.Cells[0,G.Row]+'"');

  G.DeleteRow(G.Row);
  if G.RowCount=1 then begin
    G.RowCount := G.RowCount+1;
    G.FixedRows := 1;
    G.Row := G.RowCount-1;
  end;

  if not(DS.State in [dsInsert, dsEdit]) then DS.Edit;
end;

procedure TOM_BUDJAL.BUpClick(Sender: TObject);
var
  G : THGrid;
begin
       if Ecran.ActiveControl = FListeG then G := FListeG
  else if Ecran.ActiveControl = FListeS then G := FListeS
                                        else Exit;

  Modifier := True;

  if G.Row=1 then begin
    G.SetFocus;
    Exit;
  end;
  G.ExchangeRow(G.Row,G.Row-1);
  G.Row := G.Row-1;
  G.SetFocus;
  if not(DS.State in [dsInsert, dsEdit]) then DS.Edit;
end;

procedure TOM_BUDJAL.BDownClick(Sender: TObject);
var
  G : THGrid;
begin
       if Ecran.ActiveControl = FListeG then G := FListeG
  else if Ecran.ActiveControl = FListeS then G := FListeS
                                        else Exit;

  Modifier := True;

  if G.Row=G.RowCount-1 then begin
    G.SetFocus;
    Exit;
  end;
  G.ExchangeRow(G.Row,G.Row+1);
  G.Row := G.Row+1;
  G.SetFocus;
  if not(DS.State in [dsInsert, dsEdit]) then DS.Edit;
end;

function TOM_BUDJAL.CpteAttenteOk(FicBase: TFichierBase): Boolean;
var
  St : String;
  C  : TWinControl;
begin
  Result := False;
  C := nil;
  Case FicBase of
    fbBudGen : begin
      St := GetField('BJ_GENEATTENTE');
      C := TWinControl(GetControl('BJ_GENEATTENTE',True));
    end;
    fbBudSec1..fbBudSec5 : begin
      St := GetField('BJ_SECTATTENTE');
      C := TWinControl(GetControl('BJ_SECTATTENTE',True));
    end;
  end;

  if St='' then begin
    Fiche.Pages.ActivePageIndex := 0;
    C.SetFocus;
    if FicBase=fbBudGen then MsgBox.Execute(10,'','')  // Vous devez renseigner le compte d''attente.
                        else MsgBox.Execute(14,'',''); // Vous devez renseigner la section d''attente.
    exit;
  end;

  if not LookupValueExist(TControl(C)) then begin
    Fiche.Pages.ActivePageIndex := 0;
    C.SetFocus;
    if FicBase=fbBudGen then MsgBox.Execute(11,'','')   // Le compte d''attente renseigné n''existe pas.
                        else MsgBox.Execute(15,'','') ; // La section d''attente renseignée n''existe pas.
    exit;
  end;
  Result := True;
end;

procedure TOM_BUDJAL.GriseDegriseControl;
begin
  SetControlEnabled('BJ_AXE', Not IsMouvementer);
  SetControlEnabled('BJ_COMPTEURNORMAL', Not IsMouvementer);
  SetControlEnabled('GroupBox3', Not IsMouvementer);
  SetControlEnabled('GroupBox1', Not IsMouvementer);
  SetControlEnabled('BCopieComplement', Not IsMouvementer);
end;

function TOM_BUDJAL.DoublonExiste(St: String; G: THGrid): Boolean;
var
  i : Integer ;
  StTemp : String;
begin
  Result := False;
  if G.Cells[0,1]='' then Exit;
  if G=FListeS then StTemp := Copy(St,3,Length(St))
               else StTemp := St;

  for i:=1 to G.RowCount-1 do begin
    if G.Cells[0,i]=StTemp then begin
      Result := True;
      Exit;
    end;
  end;
end;

procedure TOM_BUDJAL.BJ_EXODEBChange(Sender: TObject);
begin
  {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
  if BJ_EXODEB.Value > BJ_EXOFIN.Value then BJ_EXOFIN.Value := BJ_EXODEB.Value;
  ListePeriode(BJ_EXODEB.Value, BJPERDEB.Items, BJPERDEB.Values, True);
  PositionnePerDeb;
end;

procedure TOM_BUDJAL.BJ_EXOFINChange(Sender: TObject);
begin
  {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
  if BJ_EXOFIN.Value < BJ_EXODEB.Value then BJ_EXODEB.Value := BJ_EXOFIN.Value;
  ListePeriode(BJ_EXOFIN.Value, BJPERFIN.Items, BJPERFIN.Values, False);
  PositionnePerFin;
end;

procedure TOM_BUDJAL.BJ_AXEChange(Sender: TObject);
var
  UnFb : TFichierBase;
  szAxeValue : String;
begin
  szAxeValue := GetControlText('BJ_AXE');
  Lefb := AxeTofbBud(szAxeValue);
  UnFb := AxeToFb(szAxeValue);

  Case Lefb of
    fbBudsec1 : begin SetControlProperty('BJ_SECTATTENTE', 'DATATYPE', 'tzBudSecAtt1'); SetControlProperty('BJ_CATEGORIE', 'DATATYPE', 'ttCatJalBud1'); end;
    fbBudsec2 : begin SetControlProperty('BJ_SECTATTENTE', 'DATATYPE', 'tzBudSecAtt2'); SetControlProperty('BJ_CATEGORIE', 'DATATYPE', 'ttCatJalBud2'); end;
    fbBudsec3 : begin SetControlProperty('BJ_SECTATTENTE', 'DATATYPE', 'tzBudSecAtt3'); SetControlProperty('BJ_CATEGORIE', 'DATATYPE', 'ttCatJalBud3'); end;
    fbBudsec4 : begin SetControlProperty('BJ_SECTATTENTE', 'DATATYPE', 'tzBudSecAtt4'); SetControlProperty('BJ_CATEGORIE', 'DATATYPE', 'ttCatJalBud4'); end;
    fbBudsec5 : begin SetControlProperty('BJ_SECTATTENTE', 'DATATYPE', 'tzBudSecAtt5'); SetControlProperty('BJ_CATEGORIE', 'DATATYPE', 'ttCatJalBud5'); end;
  end;

  if Mode<>taConsult then begin
    SetControlEnabled('TBJ_Categorie', VH^.Cpta[UnFb].Structure);
    BJ_Categorie.Enabled := VH^.Cpta[UnFb].Structure;
    SetControlEnabled('TBJ_SousPlan', VH^.Cpta[UnFb].Structure);
    BJ_SousPlan.Enabled := VH^.Cpta[UnFb].Structure;
    SetControlEnabled('BZoomSousPlan', VH^.Cpta[UnFb].Structure);
  end;
end;

procedure TOM_BUDJAL.BJ_AXEClick(Sender: TObject);
begin
  if DS.State<>dsInsert then begin
    SetField('BJ_BUDSECTS', '');
    SetField('BJ_BUDSECTS2', '');
    FListeS.VidePile(False);
  end;
end;

procedure TOM_BUDJAL.BJ_PERDEBChange(Sender: TObject);
begin
  {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
//  if DS.State<>dsInsert then exit;
  SetField('BJ_PERDEB', VarToDateTime(BJPERDEB.Value));
  if ((BJPERDEB.Value <> '') and (BJPERFIN.Value = '')) or
     (VarToDateTime(BJPERDEB.Value) > VarToDateTime(BJPERFIN.Value)) then begin
    BJPERFIN.Value := DateToStr(FinDeMois(StrToDate(BJPERDEB.Value)));
    SetField('BJ_PERFIN', VarToDateTime(BJPERFIN.Value));
  end;
end;

procedure TOM_BUDJAL.BJ_PERFINChange(Sender: TObject);
begin
  {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
//  if DS.State<>dsInsert then exit;
  SetField('BJ_PERFIN', VarToDateTime(BJPERFIN.Value));
  if ((BJPERFIN.Value <> '') and (BJPERDEB.Value = '')) or
     (VarToDateTime(BJPERFIN.Value) < VarToDateTime(BJPERDEB.Value)) then begin
    BJPERDEB.Value := DateToStr(DebutDeMois(StrToDate(BJPERFIN.Value)));
    SetField('BJ_PERDEB', VarToDateTime(BJPERDEB.Value));
  end;
end;

function TOM_BUDJAL.ExoOuv: Boolean;
var
  QLoc : TQuery;
begin
  Result := False;
  QLoc := OpenSQL('SELECT EX_EXERCICE FROM EXERCICE WHERE EX_ETATBUDGET="OUV"',True);
  if QLoc.Eof then begin
    Ferme(QLoc);
    Exit;
  end;

  while Not QLoc.Eof do begin
    if QLoc.Fields[0].AsString = GetField('BJ_EXODEB') then begin
      Ferme(QLoc);
      Result := True;
      Exit;
    end;
    QLoc.Next;
  end;
  Ferme(QLoc);
end;

procedure TOM_BUDJAL.PositionnePerDeb;
begin
    {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
  if not(DS.State in [dsInsert]) then exit;
  if BJPERDEB.Items.Count > 0 then BJPERDEB.ItemIndex := 0;
  if BJPERDEB.Value <> '' then SetField('BJ_PERDEB', StrToDate(BJPERDEB.Value));
end;

procedure TOM_BUDJAL.PositionnePerFin;
begin
    {JP 25/08/05 : FQ 16514 : Enregistrement des périodes de début et de fin}
  if not(DS.State in [dsInsert]) then exit;
  BJPERFIN.ItemIndex := BJPERFIN.Items.Count-1;
  if BJPERFIN.Value <> '' then SetField('BJ_PERFIN', StrToDate(BJPERFIN.Value));
end;

procedure TOM_BUDJAL.GbCpteEnter(Sender: TObject);
begin
  if not (Ecran.ActiveControl = FListeG) and
     not (Ecran.ActiveControl = FListeS) then
    FListeG.SetFocus;
end;

procedure TOM_BUDJAL.BJ_NATJALChange(Sender: TObject);
begin
  Case BJ_NATJAL.Value[1] of
    'C' : SetField('BJ_SENS', 'D');
    'P' : SetField('BJ_SENS', 'C');
  end;
end;

function TOM_BUDJAL.CoherenceSection: Boolean;
var
  i : Integer;
  QLoc : TQuery;
  SQL : String;
begin
  Result := False;
  if FListeS.Cells[0,FListeS.RowCount-1]='' then begin
    Result := True;
    Exit;
  end;

  for i:=1 to FListeS.RowCount-1 do begin
    SQL := 'SELECT BS_BUDSECT FROM BUDSECT WHERE BS_BUDSECT="'+FListeS.Cells[0,i]+'" AND BS_AXE="'+ GetControlText('BJ_AXE')+'"';
    QLoc := OpenSQL(SQL, True);
    if QLoc.Eof then begin
      Ferme(QLoc);
      Fiche.Pages.ActivePageIndex := 1;
      if FListeS.CanFocus then FListeS.Focused;
      MsgBox.Execute(16,'',''); // Des sections n''appartiennent pas à l''axe choisi.
      Exit;
    end;
  end;
  Ferme(QLoc);
  Result := True;
end;

procedure TOM_BUDJAL.BZoomSousPlanClick(Sender: TObject);
var
  LaCatBud : TUneCatBud;
  St,Lib,Abr : String;
begin
  SetFocusControl('BJ_SOUSPLAN');
  LaCatBud := QuelleCatBud(BJ_CATEGORIE.Value);
  if LaCatBud.Code<>'' then begin
    St := ChoisirSousSectionBudget(LaCatBud, BJ_SOUSPLAN.Text);
    if (St<>BJ_SOUSPLAN.Text) And (St<>'') then begin
      if DS.State=dsBrowse then DS.Edit;
      SetField('BJ_SOUSPLAN', St);
      LibelleCatBud(GetField('BJ_CATEGORIE'), GetField('BJ_SOUSPLAN'), Lib, Abr);
      SetField('BJ_LIBELLE', Lib);
      SetField('BJ_ABREGE', Abr);
    end;
  end;
end;

procedure TOM_BUDJAL.BJ_SOUSPLANKeyPress(Sender: TObject; var Key: Char);
begin
  if key<>#8 then Key := #0;
end;

procedure TOM_BUDJAL.BCopieListe(Qui: Integer);
var
  St : String;
  Q : TQuery;
  szAxeValue, szBJ_CATEGORIE : String;
begin
  szAxeValue := GetControlText('BJ_AXE');
  If szAxeValue='' Then Exit;
  St := Choisir(MsgBox.Mess[17],'BUDJAL','BJ_LIBELLE','BJ_BUDJAL',
                'BJ_AXE="'+szAxeValue+'" AND BJ_BUDJAL<>"'+GetField('BJ_BUDJAL')+'"','BJ_BUDJAL');
  if St='' then exit;
  if MsgBox.Execute(18,'','')<>mrYes then Exit; // Confirmez-vous la recopie des comptes du budget sélectionné ?
  Modifier := True;
  ExecuteSql('DELETE FROM CROISCPT WHERE CX_JAL="'+GetField('BJ_BUDJAL')+'" And CX_TYPE="BUD"');
  if (Qui=0) Or (Qui=2) then FListeG.VidePile(False);
  if (Qui=1) Or (Qui=2) then FListeS.VidePile(False);
  szBJ_CATEGORIE := GetField('BJ_CATEGORIE');
  Q := OpenSQL('SELECT BJ_CATEGORIE,BJ_BUDGENES, BJ_BUDSECTS, BJ_BUDGENES2, BJ_BUDSECTS2 FROM BUDJAL WHERE BJ_BUDJAL="'+St+'"', True);
  if (szBJ_CATEGORIE<>'') And (szBJ_CATEGORIE<>Q.FindField('BJ_CATEGORIE').AsString) Then
    ExecuteSql('DELETE FROM CROISCPT WHERE CX_JAL="'+szBJ_CATEGORIE+'" And CX_TYPE="BUD"');
  if not Q.Eof then begin
    if (Qui=0) Or (Qui=2) then begin
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES').AsString), FListeG, True);
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES2').AsString), FListeG, False);
    end;
    if (Qui=1) Or (Qui=2) then begin
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS').AsString), FListeS, True);
      LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS2').AsString), FListeS, False);
    end;
  end;
  Ferme(Q);
  if (Qui=0) Or (Qui=2) then begin
    if FListeG.RowCount=1 then begin
      FListeG.RowCount := FListeG.RowCount+1;
      FListeG.FixedRows := 1;
      FListeG.Row := FListeG.RowCount-1;
    end;
  end;

  if (Qui=1) Or (Qui=2) then begin
    if FListeS.RowCount=1 then begin
      FListeS.RowCount := FListeS.RowCount+1;
      FListeS.FixedRows := 1;
      FListeS.Row := FListeS.RowCount-1;
    end;
  end;
end;

procedure TOM_BUDJAL.GenereCroisement(CodeACreer, CodeChoisi: String);
var
  Q : TQuery;
  T : Tob;
begin
  EnableControls(Ecran,False);
  BeginTrans;
  Q := OpenSQL('SELECT * FROM CROISCPT WHERE CX_TYPE="BUD" AND CX_JAL="'+CodeChoisi+'"', True);
  InitMove(Q.RecordCount, MsgBox.Mess[21]); // Génération des croisements budgétaires en cours...

  while Not Q.Eof do begin
    MoveCur(FALSE);
    T := Tob.Create('CROISCPT', nil, -1);
    T.InitValeurs ;
    T.PutValue('CX_TYPE', Q.FindField('CX_TYPE').AsString);
    T.PutValue('CX_JAL', CodeACreer);
    T.PutValue('CX_COMPTE', Q.FindField('CX_COMPTE').AsString);
    T.PutValue('CX_SECTION', Q.FindField('CX_SECTION').AsString);
    T.PutValue('CX_INFO', Q.FindField('CX_INFO').AsString);
    T.InsertDB(nil) ;
    T.Free ;
    Q.Next ;
  end;

  Ferme(Q);
  FiniMove;
  CommitTrans;
  EnableControls(Ecran, True);
end;

procedure TOM_BUDJAL.BCopierJalClick(Sender: TObject);
var
  St : String;
  Q : TQuery;
begin
  if Trim(GetControlText('BJ_BUDJAL'))='' then begin
    MsgBox.Execute(20,'',''); // Opération impossible : Vous n''avez pas renseigné le code identifiant du budget.
    SetFocusControl('BJ_BUDJAL');
    Exit;
  end;

  St := Choisir(MsgBox.Mess[17],'BUDJAL','BJ_LIBELLE','BJ_BUDJAL','BJ_BUDJAL<>"'+GetField('BJ_BUDJAL')+'"','BJ_BUDJAL');
  if St='' then exit;
  if MsgBox.Execute(19,'','')<>mrYes then Exit; // Confirmez-vous la recopie du budget sélectionné ?
  FListeG.VidePile(False);
  FListeS.VidePile(False);
  Q := OpenSQL('SELECT * FROM BUDJAL WHERE BJ_BUDJAL="'+St+'"', True);
  if not Q.Eof then begin
    SetField('BJ_AXE', Q.FindField('BJ_AXE').AsString);
    SetField('BJ_NATJAL', Q.FindField('BJ_NATJAL').AsString);
    SetField('BJ_CATEGORIE', Q.FindField('BJ_CATEGORIE').AsString);
    SetField('BJ_SENS', Q.FindField('BJ_SENS').AsString);
    SetField('BJ_COMPTEURNORMAL', Q.FindField('BJ_COMPTEURNORMAL').AsString);
    SetField('BJ_GENEATTENTE', Q.FindField('BJ_GENEATTENTE').AsString);
    SetField('BJ_SECTATTENTE', Q.FindField('BJ_SECTATTENTE').AsString);
    SetField('BJ_EXODEB', Q.FindField('BJ_EXODEB').AsString);
    SetField('BJ_EXOFIN', Q.FindField('BJ_EXOFIN').AsString);
    BJ_EXODEBChange(Nil);
    BJ_EXOFinChange(Nil);

    {JP 02/08/05 : FQ 16062 : Les champs sont de type DateTime}
    SetField('BJ_PERDEB', Q.FindField('BJ_PERDEB').AsDateTime);
    SetField('BJ_PERFIN', Q.FindField('BJ_PERFIN').AsDateTime);

    LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES').AsString), FListeG, True);
    LitBudgensBudsects(Trim(Q.FindField('BJ_BUDGENES2').AsString), FListeG, False);
    LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS').AsString), FListeS, True);
    LitBudgensBudsects(Trim(Q.FindField('BJ_BUDSECTS2').AsString), FListeS, False);
    if GetField('BJ_CATEGORIE')='' then GenereCroisement(GetControlText('BJ_BUDJAL'), St);
  end;
  Ferme(Q);

  if FListeG.RowCount=1 then begin
    FListeG.RowCount := FListeG.RowCount+1;
    FListeG.FixedRows := 1;
    FListeG.Row := FListeG.RowCount-1;
  end;

  if FListeS.RowCount=1 then begin
    FListeS.RowCount := FListeS.RowCount+1;
    FListeS.FixedRows := 1;
    FListeS.Row := FListeS.RowCount-1;
  end;
end;

procedure TOM_BUDJAL.BJ_CATEGORIEChange(Sender: TObject);
begin
  SetField('BJ_SOUSPLAN', '');
end;

procedure TOM_BUDJAL.CopieCatSurJournaux;
var
  St,StW,ListeC,BudG,BudG2,BudS,BudS2 : String;
  Ind : Integer;
  First : boolean;
  szBJ_CATEGORIE, szBJ_BUDJAL : String;
begin
  szBJ_CATEGORIE := GetField('BJ_CATEGORIE');
  if szBJ_CATEGORIE='' then Exit;
  if not BGModifier and not BSModifier then Exit;
  if MsgBox.Execute(22,'','')<>mrYes then Exit; // Désirez-vous apporter les modifications des comptes / sections à d''autres budgets de la catégorie concernée ?
  StW := '';
  szBJ_BUDJAL := GetField('BJ_BUDJAL');
  StW := 'BJ_CATEGORIE="'+szBJ_CATEGORIE+'" AND BJ_BUDJAL<>"'+szBJ_BUDJAL+'"' ; ;
  ListeC := ChoisirMulti(MsgBox.Mess[23],'BUDJAL','BJ_BUDJAL','BJ_LIBELLE',StW,'BJ_BUDJAL','');
  Ind := 1;
  BudG2 := '';
  BudG := EcritBudgensBudsects(FListeG,Ind);
  if Ind>0 then BudG2 := EcritBudgensBudsects(FListeG,Ind);
  Ind := 1;
  BudS2 := '';
  BudS := EcritBudgensBudsects(FListeS,Ind);
  if Ind>0 then BudS2 := EcritBudgensBudsects(FListeS,Ind);

  St := 'UPDATE BUDJAL SET ';
  if BGModifier then St := St + 'BJ_BUDGENES="'+BudG+'",BJ_BUDGENES2="'+BudG2+'"';
  if BSModifier then begin
    if BGModifier then St := St +',';
    St := St + 'BJ_BUDSECTS="'+BudS+'",BJ_BUDSECTS2="'+BudS2+'"';
  end;
  St := St + ' WHERE BJ_BUDJAL<>"'+szBJ_BUDJAL+'" AND BJ_CATEGORIE="'+szBJ_CATEGORIE+'"';
  StW := ReadTokenSt(ListeC);
  First := True;

  while StW<>'' do begin
    if First then St := St+' AND ('
             else St := St+' OR ';
    First := False;
    St := St + ' BJ_BUDJAL="'+StW+'"';
    StW := ReadTokenSt(ListeC);
  end;
  if not First then St := St+')';
  ExecuteSQL(St);

  BSModifier := False;
  BGModifier := False;
end;

procedure TOM_BUDJAL.BCopieCptClick(Sender: TObject);
begin
  BCopieListe(0);
end;

procedure TOM_BUDJAL.BCopieSectClick(Sender: TObject);
begin
  BCopieListe(1);
end;

procedure TOM_BUDJAL.BCopieDeuxClick(Sender: TObject);
begin
  BCopieListe(2);
end;

procedure TOM_BUDJAL.FListeGSorted(Sender: TObject);
begin
  FListeG.SortedCol := 0;
  Modifier := True;
end;

procedure TOM_BUDJAL.FListeSSorted(Sender: TObject);
begin
  FListeS.SortedCol := 0;
  Modifier := True;
end;

procedure TOM_BUDJAL.UpdateDate;
var
  szDate : String;
begin
  szDate := FormatDateTime('dd mmm yyyy', GetField('BJ_DATECREATION'));
  SetControlText('DATECREATION', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BJ_DATEMODIF'));
  SetControlText('DATEMODIF', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BJ_DATEOUVERTURE'));
  SetControlText('DATEOUVERTURE', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BJ_DATEFERMETURE'));
  SetControlText('DATEFERMETURE', szDate);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BUDJAL.ListeWidth;
{---------------------------------------------------------------------------------------}
begin
  FListeG.ColWidths[0] := 190;
  FListeS.ColWidths[0] := 160;
  FListeS.ColWidths[1] := 30;
end;

{JP 24/08/05 : FQ 16003 : lors de l'affichage de la fiche les colonnes des grilles
               sont redimensionnées sur la base de la taille de la fiche non sur
               celle de la grille
{---------------------------------------------------------------------------------------}
procedure TOM_BUDJAL.OnEcranResize(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {FormResize hérité}
  if Assigned(OnAncetreEcranResize) then OnAncetreEcranResize(Sender);
  {Retaille les grilles}
  ListeWidth;
end;

Initialization
  registerclasses ( [ TOM_BUDJAL ] ) ;

end.
