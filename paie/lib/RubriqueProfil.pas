{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 15/06/2001
Modifié le ... : 30/07/2001
Description .. : Gestion Associé profil
Mots clefs ... : PAIE;PROFIL
*****************************************************************}
{
PT1   : 15/06/2001 SB      Déselection des rubriques : Après selection puis
                           affectation de grille à grille on déselectionne..
PT2   : 07/11/2001 PH      On ne prend plus la nature du profil comme critere
                           dans profilrub car elle ne fait plus partie de la clé
                           primaire
PT3   : 14/11/2001 SB V562 Affectation du prédefini et du nodossier erroné
PT4   : 28/11/2001 SB V563 ReCorrection mauvaise affectation champ predefini
PT5   : 28/11/2001 SB V563 Oublie ds correction PT2
PT6   : 08/01/2002 SB V571 Correction affectation pred Cotisation
PT7   : 24/04/2002 PH V582 Gestion associée protégée de la même façon que la
                           fiche qui l'appelle
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT8   : 20/06/2003 SB V_42 Ajout indication du retour de modifcation pour
                           rechargement contexte paie
PT9   : 29/09/2003 PH V_421 gestion des caches en CWAS
PT10  : 02/02/2005 PH V_602 FQ 11635 récupération du libellé de la rubrique
PT11  : 06/12/2005 SB V_65 FQ 12666 impression de la gestion associée
PT12  : 14/05/2007 MF V_72 FQ 13977 : modification gestion associée profil STD
                           impossible qd utilisatueur non réviseur
PT13  : 20/06/2007 FC V_72 FQ 14353 les lignes affichées doivent être filtrées
                           en fonction de l'activité associée au profil
PT14  : 12/09/2007 PH V_80 FQ 14419 Remis accès alimentation des profils CEGID
                           par une rubrique STD ou DOS
}
unit RubriqueProfil;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, ComCtrls,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} EdtREtat,
  {$ELSE}
  UtilEAGL,
  {$ENDIF}
  Buttons, Hent1, Grids, Hctrls, StdCtrls, Utob, HTB97, uiutil, ExtCtrls,
  hmsgbox, HPanel,
  PGoutils,PgOutils2;

function ProfilPaieDetail(PROFIL, LIBELLE, TYPEACTION: string;ACTIVITE: string = ''): boolean; //NATURE

type
  TFProfilRubrique = class(TForm)
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BVALIDER: TToolbarButton97;
    BFERMER: TToolbarButton97;
    BHELP: TToolbarButton97;
    BImprimer: TToolbarButton97;
    pages: TPageControl;
    TSCritere: TTabSheet;
    GRID_PROFILRUB: THGrid;
    Panel1: TPanel;
    P: THLabel;
    LProfil: TLabel;
    Panel2: TPanel;
    BNONRECUP_REM: THSpeedButton;
    BNONRECUP_COT: THSpeedButton;
    BNRETIRE_RUB: THSpeedButton;
    THEMEREM: TLabel;
    THEMECOT: TLabel;
    GRID_REMUNERATION: THGrid;
    GRID_COTISATION: THGrid;
    CbxThemeRem: THValComboBox;
    CbxThemeCot: THValComboBox;
    Msg: THMsgBox;
    Splitter1: TSplitter;
    FLISTE: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure CbxThemeCotChange(Sender: TObject);
    procedure CbxThemeRemChange(Sender: TObject);
    procedure DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);
    procedure TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure GRID_PROFILRUBMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BVALIDERClick(Sender: TObject);
    procedure BNONRECUP_REMClick(Sender: TObject);
    procedure BNRETIRE_RUBClick(Sender: TObject);
    procedure BNONRECUP_COTClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BHELPClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject); { PT11 }

  private
    { Déclarations privées }
    TOB_REMUNERATION: TOB;
    TOB_COTISATION: TOB;
    TOB_REMUNERATIONCACHE: TOB;
    TOB_COTISATIONCACHE: TOB;
    TOB_PROFILRUB: TOB;
    NATURE, PROFIL, LIBELLE, ACTIVITE: string;  //PT13
    Modifie: Boolean;
    Quoi: string;
    procedure AfficheGrids;
    procedure VireDejaUtilise;
    procedure Sauve;
  public
    { Déclarations publiques }
    Retour: Boolean; // PT-8
  end;


implementation
uses P5Def;
{$R *.DFM}
function ProfilPaieDetail(PROFIL, LIBELLE, TYPEACTION: string;ACTIVITE: string = ''): boolean; //NATURE
var
  X: TFProfilRubrique;
  PP: THPanel;
begin
  Result := False; // PT-8
  X := TFProfilRubrique.Create(Application);
  X.Quoi := TYPEACTION;
  {  If (TYPEACTION = 'C') AND (V_PGI.Debug = FALSE)
     then X.Quoi := FALSE;
    if X.Quoi = FALSE then
      begin
      X.BNONRECUP_REM.enabled := FALSE;
      X.BNONRECUP_COT.enabled := FALSE;
      X.BNRETIRE_RUB.enabled := FALSE;
      X.BVALIDER.enabled := FALSE;
      end;}
   // X.NATURE    := NATURE ;   PT5
  X.PROFIL := PROFIL;
  X.LIBELLE := LIBELLE;
  X.ACTIVITE := ACTIVITE; //PT13
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      X.ShowModal;
      Result := X.Retour; // PT-8
    finally
      X.Free;
    end;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

procedure TFProfilRubrique.FormShow(Sender: TObject);
var
  ST: string;
  Q: TQuery;
  LectureSeule, CEG, STD, DOS: boolean;
begin

  //  PT7 : 24/04/2002 PH Gestion associée protégée de la même façon que la fiche qui l'appelle
  AccesPredefini(QUOI, CEG, STD, DOS);
  if QUOI = 'CEG' then LectureSeule := (CEG = False)
  else if QUOI = 'STD' then LectureSeule := (STD = False)
  else LectureSeule := (DOS = False);
  if LectureSeule then
  begin
    BVALIDER.Enabled := FALSE;
    BNONRECUP_REM.Enabled := FALSE;
    BNONRECUP_COT.Enabled := FALSE;
    BNRETIRE_RUB.Enabled := FALSE;
  end;
  // PPI_NATURE.Text := NATURE ;
  LProfil.Caption := PROFIL + ' ' + LIBELLE;
  Modifie := FALSE;
  Retour := FALSE; // PT-8
  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('REMUNERATION');
  AvertirCacheServer('COTISATION');
  {$ENDIF}
  // CHARGEMENT DE LA TABLE DES REMUNERATIONS
  TOB_REMUNERATION := TOB.Create('Table des Rémunérations', nil, -1);
  //DEB PT13
  if (ACTIVITE <> '') then
    st := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## (PRM_ACTIVITE="' + ACTIVITE + '" OR PRM_ACTIVITE="")'
  else
    st := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI##';
  Q := OpenSql(st, TRUE);
  //FIN PT13
  TOB_REMUNERATION.LoadDetailDB('REMUNERATION', '', '', Q, False);
  Ferme(Q);    //PT13

  // CHARGEMENT DE LA TABLE DES COTISATIONS
  TOB_COTISATION := TOB.Create('Table des Cotisations', nil, -1);
  //DEB PT13
  if (ACTIVITE <> '') then
    st := 'SELECT * FROM COTISATION WHERE ##PCT_PREDEFINI## (PCT_ACTIVITE="' + ACTIVITE + '" OR PCT_ACTIVITE="")'
  else
    st := 'SELECT * FROM COTISATION WHERE ##PCT_PREDEFINI##';
  Q := OpenSql(st, TRUE);
  //FIN PT13
  TOB_COTISATION.LoadDetailDB('COTISATION', '', '', Q, False);
  Ferme(Q);    //PT13


  // CHARGEMENT DE LA TABLE DES RUBRIQUES DES PROFILS
  TOB_PROFILRUB := TOB.Create('Table des ProfilRubriques', nil, -1);
  //PT2 : 07/11/2001 PH  On ne prend plus la nature du profil comme critere dans profilrub
  st := 'SELECT * FROM PROFILRUB WHERE ##PPM_PREDEFINI## PPM_PROFIL="' + PROFIL + '"';
  Q := OpenSql(st, TRUE);
  TOB_PROFILRUB.LoadDetailDB('PROFILRUB', '', '', Q, False);
  FERME(Q);
  // CREATION DES TOB CACHEES
  TOB_COTISATIONCACHE := TOB.Create('Table des cotisations cachées', nil, -1);
  TOB_REMUNERATIONCACHE := TOB.Create('Table des remunerations cachées', nil, -1);

  VireDejaUtilise;
  AfficheGrids;
end;

procedure TFProfilRubrique.AfficheGrids;
begin
  TOB_REMUNERATION.PutGridDetail(GRID_REMUNERATION, FALSE, TRUE, 'PRM_RUBRIQUE;PRM_LIBELLE', TRUE);
  GRID_REMUNERATION.SortGrid(0, FALSE);

  TOB_COTISATION.PutGridDetail(GRID_COTISATION, FALSE, TRUE, 'PCT_RUBRIQUE;PCT_LIBELLE', TRUE);
  GRID_COTISATION.SortGrid(0, FALSE);

  TOB_PROFILRUB.PutGridDetail(GRID_PROFILRUB, FALSE, TRUE, 'PPM_NATURERUB;PPM_RUBRIQUE;PPM_LIBELLE;PPM_IMPRIMABLE', TRUE);
  GRID_PROFILRUB.SortGrid(0, FALSE);
end;


procedure TFProfilRubrique.CbxThemeCotChange(Sender: TObject);
var
  ThemeCot: string;
  I: integer;
  T1: TOB;
begin
  ThemeCot := CbxThemeCot.Value;
  for I := TOB_COTISATIONCACHE.Detail.Count - 1 downto 0 do
  begin
    T1 := TOB_COTISATIONCACHE.Detail[I];
    T1.ChangeParent(TOB_COTISATION, 0);
  end;

  if ThemeCot <> '' then
    for I := TOB_COTISATION.Detail.Count - 1 downto 0 do
    begin
      T1 := TOB_COTISATION.Detail[I];
      if (T1.GetValue('PCT_THEMECOT') <> ThemeCot) then
      begin
        // ON AJOUTE DANS TOB_COTISATIONCACHE
        T1.ChangeParent(TOB_COTISATIONCACHE, 0);
      end;
    end;
  VireDejaUtilise;
  AfficheGrids;
end;

procedure TFProfilRubrique.CbxThemeRemChange(Sender: TObject);
var
  ThemeRem: string;
  I: integer;
  T1: TOB;
begin
  ThemeRem := CbxThemeRem.Value;

  for I := TOB_REMUNERATIONCACHE.Detail.Count - 1 downto 0 do
  begin
    T1 := TOB_REMUNERATIONCACHE.Detail[I];
    T1.ChangeParent(TOB_REMUNERATION, 0);
  end;
  if ThemeRem <> '' then
    for I := TOB_REMUNERATION.Detail.Count - 1 downto 0 do
    begin
      T1 := TOB_REMUNERATION.Detail[I];
      if (T1.GetValue('PRM_THEMEREM') <> ThemeRem) then
      begin
        // ON AJOUTE DANS TOB_REMUNERATIONCACHE
        T1.ChangeParent(TOB_REMUNERATIONCACHE, 0);
      end;
    end;
  VireDejaUtilise;
  AfficheGrids;
end;

procedure TFProfilRubrique.TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  //
  Accept := FALSE;
  if (Sender is THgrid) then Accept := TRUE;
end;


procedure TFProfilRubrique.DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);
var
  T, TD, P: TOB;
  MultiSel: Boolean;
  Nbre, i: Integer; // recherche du nombre d'objets selectionnes
  CEG, STD, DOS, LectureSeule : Boolean; // PT12
begin
  if Sender = Source then EXIT;

// d PT12
  AccesPredefini(QUOI, CEG, STD, DOS);
  if QUOI = 'CEG' then
    LectureSeule := (CEG = False)
  else
    if QUOI = 'STD' then
      LectureSeule := (STD = False)
    else
      LectureSeule := (DOS = False);

  if LectureSeule AND (Quoi <> 'CEG') then // PT14
    EXIT;
// f PT12

  Modifie := TRUE;
  Retour := TRUE; // PT-8
  Nbre := THGrid(Source).RowCount - 1;
  MultiSel := THGrid(Source).MultiSelect;
  if not MultiSel then Nbre := 1;
  for i := 1 to Nbre do
  begin
    if THGrid(Source).IsSelected(i) then
    begin
      if (Sender = GRID_REMUNERATION) or (Sender = GRID_COTISATION) then
      begin
        if not MultiSel then TD := TOB(THGrid(Source).Objects[0, THGrid(Source).Row])
        else TD := TOB(THGrid(Source).Objects[0, i]);

        //     TD := TOB (THGrid(Source).Objects[0, THGrid(Source).Row]) ;
        if TD = nil then exit;
        if (TD.GetValue('PPM_NATURERUB') = 'AAA') then
        begin
          T := TOB_REMUNERATIONCACHE.FindFirst(['PRM_RUBRIQUE'], [TD.GetValue('PPM_RUBRIQUE')], FALSE);
          if (Quoi = 'CEG') and (T.getvalue('PRM_PREDEFINI') = 'CEG') and (T <> nil) and (V_PGI.Debug = FALSE) then
          begin
            PGIBox('Vous ne pouvez pas enlever une rubrique CEGID d''un profil CEGID', 'Gestion des profils');
            exit;
          end;
          if T <> nil then T.ChangeParent(TOB_REMUNERATION, -1) else ShowMessage('Pas trouvé');
          TD.Free;
        end
        else
        begin
          T := TOB_COTISATIONCACHE.FindFirst(['PCT_RUBRIQUE'], [TD.GetValue('PPM_RUBRIQUE')], FALSE);
          if (Quoi = 'CEG') and (T.getvalue('PCT_PREDEFINI') = 'CEG') and (T <> nil) and (V_PGI.Debug = FALSE) then
          begin
            PGIBox('Vous ne pouvez pas enlever une rubrique CEGID d''un profil CEGID', 'Gestion des profils');
            exit;
          end;
          if T <> nil then T.ChangeParent(TOB_COTISATION, -1) else ShowMessage('Pas trouvé');
          TD.Free;
        end;

      end
      else
      begin
        // TD AU FORMAT PROFILRUB
        if not MultiSel then TD := TOB(THGrid(Source).Objects[0, THGrid(Source).Row])
        else TD := TOB(THGrid(Source).Objects[0, i]);

        //     TD := TOB (THGrid(Source).Objects[0, THGrid(Source).Row]) ;
        if TD = nil then exit;
        P := TOB_PROFILRUB;
        if (Source = GRID_REMUNERATION) or (Source = GRID_COTISATION) then
        begin
          // DANS LE CAS OU LA SOURCE = GRID_REMUNERATION
          if Source = GRID_REMUNERATION then
          begin
            if (Quoi = 'CEG') and (TD.GetValue('PRM_PREDEFINI') = 'CEG') and (V_PGI.Debug = FALSE) then
            begin
              PGIBox('Vous ne pouvez pas rajouter une rubrique CEGID dans un profil CEGID', 'Gestion des profils');
              exit;
            end;
            T := TOB.Create('PROFILRUB', P, -1);
            TD.ChangeParent(TOB_REMUNERATIONCACHE, -1);
            T.PutValue('PPM_NATURERUB', TD.GetValue('PRM_NATURERUB'));
            T.PutValue('PPM_RUBRIQUE', TD.GetValue('PRM_RUBRIQUE'));
            T.PutValue('PPM_LIBELLE', TD.GetValue('PRM_LIBELLE'));
            T.PutValue('PPM_IMPRIMABLE', TD.GetValue('PRM_IMPRIMABLE'));
            T.PutValue('PPM_PROFIL', PROFIL);
            T.PutValue('PPM_TYPEPROFIL', NATURE);
            //DEB PT3
            T.PutValue('PPM_PREDEFINI', TD.GetValue('PRM_PREDEFINI'));
            T.PutValue('PPM_NODOSSIER', TD.GetValue('PRM_NODOSSIER'));
            if (Quoi = 'STD') and (TD.GetValue('PRM_PREDEFINI') = 'CEG') then //PT3
              T.PutValue('PPM_PREDEFINI', 'STD'); //PT3
            if Quoi = 'DOS' then
            begin
              T.PutValue('PPM_PREDEFINI', 'DOS');
              // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
              T.PutValue('PPM_NODOSSIER', PgRendNoDossier());
            end;
            {if Quoi = 'CEG' then
             T.PutValue('PPM_PREDEFINI',TD.GetValue('PRM_PREDEFINI'))
             else T.PutValue('PPM_PREDEFINI',Quoi) ;
            if (Quoi='CEG') OR (Quoi='STD') then T.PutValue('PPM_NODOSSIER',TD.GetValue('PRM_NODOSSIER'))
             else
             begin
             if (V_PGI_Env <> NIL) then T.PutValue('PPM_NODOSSIER',V_PGI_env.nodossier)
              else T.PutValue('PPM_NODOSSIER','000000');
             end;}
            //FIN PT3
          end
          else
          begin
            if (Quoi = 'CEG') and (TD.GetValue('PCT_PREDEFINI') = 'CEG') and (V_PGI.Debug = FALSE) then
            begin
              PGIBox('Vous ne pouvez pas rajouter une rubrique CEGID dans un profil CEGID', 'Gestion des profils');
              exit;
            end;
            T := TOB.Create('PROFILRUB', P, -1);
            TD.ChangeParent(TOB_COTISATIONCACHE, -1);
            T.PutValue('PPM_NATURERUB', TD.GetValue('PCT_NATURERUB'));
            T.PutValue('PPM_RUBRIQUE', TD.GetValue('PCT_RUBRIQUE'));
            T.PutValue('PPM_LIBELLE', TD.GetValue('PCT_LIBELLE'));
            T.PutValue('PPM_IMPRIMABLE', TD.GetValue('PCT_IMPRIMABLE'));
            T.PutValue('PPM_PROFIL', PROFIL);
            T.PutValue('PPM_TYPEPROFIL', NATURE);
            { DEB PT3
            if Quoi = 'CEG' then T.PutValue('PPM_PREDEFINI',TD.GetValue('PCT_PREDEFINI'))
             else T.PutValue('PPM_PREDEFINI',Quoi) ;
            if (Quoi='CEG') OR (Quoi='STD') then T.PutValue('PPM_NODOSSIER',TD.GetValue('PCT_NODOSSIER'))
             else
             begin
             if (V_PGI_Env <> NIL) then T.PutValue('PPM_NODOSSIER',V_PGI_env.nodossier)
              else T.PutValue('PPM_NODOSSIER','000000');
             end;  }
            T.PutValue('PPM_PREDEFINI', TD.GetValue('PCT_PREDEFINI'));
            T.PutValue('PPM_NODOSSIER', TD.GetValue('PCT_NODOSSIER'));
            if (Quoi = 'STD') and (TD.GetValue('PCT_PREDEFINI') = 'CEG') then //PT6
              T.PutValue('PPM_PREDEFINI', 'STD'); //PT6
            if Quoi = 'DOS' then
            begin
              T.PutValue('PPM_PREDEFINI', 'DOS');
              // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
              T.PutValue('PPM_NODOSSIER', PgRendNoDossier());
            end;
            //FIN PT3
          end;

        end
        else
        begin

          // ON CHANGE LE PARENT
          TD.ChangeParent(P, -1);

        end;
      end;
    end;
  end; // Fin multi selection
  {DEB PT- 1}
  THGrid(Source).ClearSelected;
  {FIN PT- 1}
  AfficheGrids;
end;

procedure TFProfilRubrique.GRID_PROFILRUBMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbLeft then
  begin
    if Sender = GRID_COTISATION then
      GRID_COTISATION.BeginDrag(TRUE, 10)
    else if Sender = GRID_REMUNERATION then
      GRID_REMUNERATION.BeginDrag(TRUE, 10)
    else if Sender = GRID_PROFILRUB then
      GRID_PROFILRUB.BeginDrag(TRUE, 10);
  end;
end;


procedure TFProfilRubrique.VireDejaUtilise;
var
  I: integer;
  T, TT: TOB;
begin
  for I := TOB_PROFILRUB.Detail.Count - 1 downto 0 do
  begin
    T := TOB_PROFILRUB.Detail[I];
    TT := TOB_COTISATION.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [T.GetValue('PPM_RUBRIQUE'), T.GetValue('PPM_NATURERUB')], FALSE);
    if (TT <> nil) then
    begin
      // ON AJOUTE DANS TOB_COTISATIONCACHE
      // PT10 Récup du libellé de la rubrique
      if (T.GetValue ('PPM_LIBELLE') <> TT.GetValue ('PCT_LIBELLE')) then T.PutValue ('PPM_LIBELLE',TT.GetValue ('PCT_LIBELLE'));
      TT.ChangeParent(TOB_COTISATIONCACHE, 0);
    end;
  end;

  for I := TOB_PROFILRUB.Detail.Count - 1 downto 0 do
  begin
    T := TOB_PROFILRUB.Detail[I];
    TT := TOB_REMUNERATION.FindFirst(['PRM_RUBRIQUE', 'PRM_NATURERUB'], [T.GetValue('PPM_RUBRIQUE'), T.GetValue('PPM_NATURERUB')], FALSE);
    if (TT <> nil) then
    begin
      // ON AJOUTE DANS TOB_REMUNERATIONCACHE
      // PT10 Récup du libellé de la rubrique      
      if (T.GetValue ('PPM_LIBELLE') <> TT.GetValue ('PRM_LIBELLE')) then T.PutValue ('PPM_LIBELLE',TT.GetValue ('PRM_LIBELLE'));
      TT.ChangeParent(TOB_REMUNERATIONCACHE, 0);
    end;
  end;
end;

procedure TFProfilRubrique.BVALIDERClick(Sender: TObject);
begin
  Sauve;
  ChargementTablette('PPM', '');
  ModalResult := mrOk;
end;

procedure TFProfilRubrique.BNONRECUP_REMClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_PROFILRUB, GRID_REMUNERATION, 0, 0)
end;

procedure TFProfilRubrique.BNRETIRE_RUBClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_COTISATION, GRID_PROFILRUB, 0, 0);
end;

procedure TFProfilRubrique.BNONRECUP_COTClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_PROFILRUB, GRID_COTISATION, 0, 0)
end;

procedure TFProfilRubrique.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  Rep: integer;
begin
  if Modifie then Rep := Msg.Execute(0, caption, '') else rep := mrNo;
  case Rep of
    mrYes: Sauve;
    mrNo: Modifie := FALSE;
    mrCancel: CanClose := FALSE;
  end;
end;

procedure TFProfilRubrique.Sauve;
begin
  if not Modifie then exit;
  //PT2 : 07/11/2001 PH  On ne prend plus la nature du profil comme critere dans profilrub
  ExecuteSQL('DELETE FROM PROFILRUB WHERE ##PPM_PREDEFINI## PPM_PROFIL="' + Profil + '"');
  TOB_ProfilRub.SetAllModifie(TRUE);
  TOB_ProfilRub.InsertDB(nil, TRUE);
  Modifie := FALSE;
end;

procedure TFProfilRubrique.BHELPClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;
{ DEB PT11 }
procedure TFProfilRubrique.BImprimerClick(Sender: TObject);
Var
  St    : String;
begin
 St := ' AND PPI_PROFIL="'+PROFIL+'" ';
 LanceEtat('E', 'PGA', 'PPF', True, False, False, Pages, St, '', False);
end;
{ FIN PT11 }

end.

