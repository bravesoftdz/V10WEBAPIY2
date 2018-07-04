{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 15/06/2001
Modifié le ... : 19/06/2001
Description .. : Gestion associé des cumuls alimentant les rubriques
Suite ........ : de cotisations et de rémunérations
Mots clefs ... : PAIE;GESTIONASSOCIEE
*****************************************************************}
{PT- 1 : 15/06/2001 : SB : Déselection des cumuls
     Après selection puis affectation de grille à grille on déselectionne..
 PT-2 : 24/08/2001 : PH : Correction mauvaise affectation champ predefini
 PT-3 : 13/11/2001 SB Correction mauvaise affectation champ predefini
 PT-4   28/11/2001 SB  ReCorrection mauvaise affectation champ predefini
 PT-5   30/11/2001 SB Fiche de bug n°369 affectation d'un cumul STD à une rubrique DOS
 PT-6   26/03/2002 SB Fiche de bug n°10052 Desaffectation d'un cumul STD à une rubrique DOS
 PT-7 : 24/04/2002 PH V582 Gestion associée protégée de la même façon que la fiche qui l'appelle
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
 PT8   20/06/2003 SB V_42 Ajout indication du retour de modifcation pour rechargement contexte paie
 PT9   29/09/2003 PH V_421 gestion des caches en CWAS
 PT10  01/06/2004 PH V_50  FQ 11141 Design et ergonomie des grilles
 PT11  25/08/2004 PH V_50  FQ 11186 Controle de la modification par le glisser déplacer
 PT12  04/01/2005 PH V_60  FQ 11878 Suppression contrôle glisser déplacer
 PT13  08/08/2005 PH V_50  FQ 11186 Suite : controle si Réviseur pour laisser l'accès si STD
 PT14  13/09/2005 PH V_50  FQ 11966 Modification rubrique CEGID si non réviseur interdit 
}
unit GestionAssocie;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  {$IFNDEF EAGLCLIENT}
  Db,
  DBGrids,
  HDB,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF }
  HTB97,
  Grids,
  ExtCtrls,
  HPanel,
  utob,
  StdCtrls,
  Hctrls,
  hent1,
  uiutil,
  HSysMenu,
  Buttons,
  hmsgbox;


// TYPEACTION pour indiquer si la rubrique est predefini CEGID
function LanceGestionAssocie(NATURE, RUBRIQUE, LIBELLE, TYPEACTION: string): boolean;

type
  TFCumulGestionAssociee = class(TForm)
    HPanel1: THPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    BNONAFFECTE_GAIN: THSpeedButton;
    BGAIN_NONAFFECTE: THSpeedButton;
    BGAIN_RETENU: THSpeedButton;
    BRETENU_GAIN: THSpeedButton;
    BNONAFFECTE_RETENU: THSpeedButton;
    BRETENU_NONAFFECTE: THSpeedButton;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BVALIDER: TToolbarButton97;
    BFERMER: TToolbarButton97;
    BHELP: TToolbarButton97;
    GRID_GAIN: THGrid;
    GRID_CUMULS: THGrid;
    GRID_RETENU: THGrid;
    Msg: THMsgBox;
    LRUBRIQUE: THLabel;
    procedure FormShow(Sender: TObject);
    procedure TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);
    procedure GRID_CUMULSMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BVALIDERClick(Sender: TObject);
    procedure BNONAFFECTE_GAINClick(Sender: TObject);
    procedure BNONAFFECTE_RETENUClick(Sender: TObject);
    procedure BGAIN_NONAFFECTEClick(Sender: TObject);
    procedure BRETENU_NONAFFECTEClick(Sender: TObject);
    procedure BGAIN_RETENUClick(Sender: TObject);
    procedure BRETENU_GAINClick(Sender: TObject);
    procedure GRID_CUMULSDblClick(Sender: TObject);
    procedure GRID_GAINDblClick(Sender: TObject);
    procedure GRID_RETENUDblClick(Sender: TObject);
    procedure BHELPClick(Sender: TObject);
  private { Déclarations privées }
    Modifie: Boolean;
    LectureSeule, CEG, STD, DOS: Boolean; // autorisation des acces
    Quoi: string;
    procedure VireDejaUtilise;
    procedure AfficheGrids;
    procedure sauve;
  public
    { Déclarations publiques }
    TOB_CUMULS: TOB;
    TOB_CUMULCACHE: TOB;
    TOB_CUGAIN: TOB;
    TOB_CURETENU: TOB;
    TOB_CULIBRE: TOB;
    TOB_ANCIENCURUB: TOB;
    TOB_CURUB: TOB;
    NATURE, RUBRIQUE, LIBELLE: string;
    Retour: Boolean; //PT-8
  end;

implementation

uses P5Def, pgoutils,PgOutils2;
{$R *.DFM}

function LanceGestionAssocie(NATURE, RUBRIQUE, LIBELLE, TYPEACTION: string): boolean;
var
  X: TFCumulGestionAssociee;
  PP: THPanel;
begin
  Result := False; //PT-8
  X := TFCumulGestionAssociee.Create(Application);
  X.Quoi := TYPEACTION;
  {  // toujours accessible mais controle si le cumul est cegid ou non idem rubriqueprofil
  if (X.Quoi = 'C') AND (V_PGI.Debug = FALSE) then
    begin
    X.BNONAFFECTE_GAIN.enabled := FALSE;
    X.BGAIN_NONAFFECTE.enabled := FALSE;
    X.BGAIN_RETENU.enabled := FALSE;
    X.BRETENU_GAIN.enabled := FALSE;
    X.BNONAFFECTE_RETENU.enabled := FALSE;
    X.BRETENU_NONAFFECTE.enabled := FALSE;
    X.BVALIDER.enabled := FALSE;
    end;
    }
  X.NATURE := NATURE;
  X.RUBRIQUE := RUBRIQUE;
  X.LIBELLE := LIBELLE;
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      X.ShowModal;
      result := X.Retour; //PT-8
    finally
      X.Free;
    end;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

procedure TFCumulGestionAssociee.FormShow(Sender: TObject);
var
  ST: string;
  Q: TQuery;
  TypeP: string;
//  CEG, STD, DOS: boolean;   PT11 Doublon avec les mêmes variables de l'objet
  HMTrad: THSystemMenu;
begin
  if Quoi = 'C' then TypeP := 'CEG'
  else if Quoi = 'S' then TypeP := 'STD'
  else TypeP := 'DOS';
  //  PT-7 : 24/04/2002 PH Gestion associée protégée de la même façon que la fiche qui l'appelle
  AccesPredefini(TypeP, CEG, STD, DOS);
  if TypeP = 'CEG' then LectureSeule := (CEG = False)
  else if TypeP = 'STD' then LectureSeule := (STD = False)
  else LectureSeule := (DOS = False);
  if LectureSeule then
  begin
    BNONAFFECTE_GAIN.Enabled := FALSE;
    BGAIN_NONAFFECTE.Enabled := FALSE;
    BGAIN_RETENU.Enabled := FALSE;
    BRETENU_GAIN.Enabled := FALSE;
    BVALIDER.Enabled := FALSE;
    BNONAFFECTE_RETENU.Enabled := FALSE;
    BRETENU_NONAFFECTE.Enabled := FALSE;
  end;

  LRUBRIQUE.Caption := RUBRIQUE + '   ' + LIBELLE;
  Modifie := FALSE;
  Retour := FALSE; //PT-8
  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('CUMULPAIE');
  {$ENDIF}
  // CHARGEMENT DE LA TABLE DES CUMULS
  TOB_CUMULS := TOB.Create('Table des cumuls', nil, -1);
  TOB_CUMULS.LoadDetailDB('CUMULPAIE', '', '', nil, False);

  TOB_CUMULCACHE := TOB.Create('Table des cumuls cachés', nil, -1);

  // CHARGEMENT DE LA TABLE DES CUMULS / RUBRIQUES
  TOB_CURUB := TOB.Create('Tables des cumuls par rubrique', nil, -1);
  st := 'SELECT * FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="' + NATURE + '" AND ##PCR_PREDEFINI## PCR_RUBRIQUE="' + RUBRIQUE + '"'; //**//
  Q := OpenSql(st, TRUE);
  TOB_CURUB.LoadDetailDB('CUMULRUBRIQUE', '', '', Q, FALSE);
  Ferme(Q);

  TOB_ANCIENCURUB := TOB.Create('tables ANCIEN', nil, -1);
  TOB_ANCIENCURUB.Dupliquer(TOB_CURUB, TRUE, TRUE);

  // CREATION DE LA TOB DES CUMULS AFFECTES EN GAIN
  TOB_CUGAIN := TOB.Create('liste gain', nil, -1);

  // CREATION DE LA TOB DES CUMULS AFFECTES EN RETENU
  TOB_CURETENU := TOB.Create('liste retenu', nil, -1);

  VireDejaUtilise;
// PT10  01/06/2004 PH V_50  FQ 11141 Design et ergonomie des grilles  
  HMTrad.ResizeGridColumns(GRID_CUMULS);
  HMTrad.ResizeGridColumns(GRID_GAIN);
  HMTrad.ResizeGridColumns(GRID_RETENU);
  AfficheGrids;
end;

procedure TFCumulGestionAssociee.AfficheGrids;
begin
  TOB_CUMULS.PutGridDetail(GRID_CUMULS, FALSE, TRUE, 'PCL_CUMULPAIE;PCL_LIBELLE', TRUE);
  GRID_CUMULS.SortGrid(0, FALSE);

  TOB_CUGAIN.PutGridDetail(GRID_GAIN, FALSE, TRUE, 'PCR_CUMULPAIE;PCR_LIBELLE', TRUE);
  GRID_GAIN.SortGrid(0, FALSE);

  TOB_CURETENU.PutGridDetail(GRID_RETENU, FALSE, TRUE, 'PCR_CUMULPAIE;PCR_LIBELLE', TRUE);
  GRID_RETENU.SortGrid(0, FALSE);
end;

procedure TFCumulGestionAssociee.TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := FALSE;
  if (Sender is THgrid) then Accept := TRUE;
end;

procedure TFCumulGestionAssociee.DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);
var
  T, TD, P: TOB;
  MultiSel: Boolean;
  TypeP: string;
  Nbre, i: Integer; // recherche du nombre d'objets selectionnes
begin
  if Sender = Source then EXIT;
  // DEB PT11
  // PT12 Suppression PT11
  if LectureSeule AND (Quoi = 'S') AND (STD = FALSE) then // PT13
  begin
    PGIBox ('Vous n''êtes pas autorisé à modifier la gestion associée', 'Gestion associée');
    exit;
  end;
  // FIN PT11
  // DEB PT14
  if LectureSeule AND (Quoi = 'C') AND (STD = FALSE) then // PT13
  begin
    PGIBox ('Vous n''êtes pas autorisé à modifier la gestion associée', 'Gestion associée');
    exit;
  end;
  // FIN PT14
  Nbre := THGrid(Source).RowCount - 1;
  if Quoi = 'C' then TypeP := 'CEG'
  else if Quoi = 'S' then TypeP := 'STD'
  else TypeP := 'DOS';
  Modifie := TRUE;
  Retour := TRUE; //PT-8
  MultiSel := THGrid(Source).MultiSelect;
  if not MultiSel then Nbre := 1;
  for i := 1 to Nbre do
  begin
    if THGrid(Source).IsSelected(i) then
    begin
      if Sender = GRID_CUMULS then
      begin
        if not MultiSel then TD := TOB(THGrid(Source).Objects[0, THGrid(Source).Row])
        else TD := TOB(THGrid(Source).Objects[0, i]);
        if TD = nil then exit;

        T := TOB_CUMULCACHE.FindFirst(['PCL_CUMULPAIE'], [TD.GetValue('PCR_CUMULPAIE')], FALSE);
        if (Quoi = 'C') and (T.getvalue('PCL_PREDEFINI') = 'CEG') and (T <> nil) and (V_PGI.Debug = FALSE) then
        begin
          PGIBox('Vous ne pouvez pas enlever un cumul CEGID d''une rubrique CEGID', 'Gestion associée');
          exit;
        end;
        {PT-6 mise en commentaire
        if (DOS) AND (T.getvalue ('PCL_PREDEFINI')='STD') AND (T <> NIL) then
          begin
          PGIBox ('Vous ne pouvez pas enlever un cumul standard d''une rubrique', 'Gestion associée');
          exit;
          end;                   }

        if T <> nil then T.ChangeParent(TOB_CUMULS, -1) else ShowMessage('Pas trouvé');

        TD.Free;
      end else
      begin
        // TD AU FORMAT CUMULPAIE
        if not MultiSel then TD := TOB(THGrid(Source).Objects[0, THGrid(Source).Row])
        else TD := TOB(THGrid(Source).Objects[0, i]);

        //     TD := TOB (THGrid(Source).Objects[0, THGrid(Source).Row]) ;
        if TD = nil then exit;
        if Sender = GRID_GAIN then P := TOB_CUGAIN else P := TOB_CURETENU;

        if Source = GRID_CUMULS then
        begin
          if (Quoi = 'C') and (TD.GetValue('PCL_PREDEFINI') = 'CEG') and (V_PGI.Debug = FALSE) then
          begin
            PGIBox('Vous ne pouvez pas rajouter un cumul CEGID dans une rubrique CEGID', 'Gestion associée');
            exit;
          end;
          { PT-5 Mise en commentaire du contrôle
          if (DOS) AND (TD.GetValue ('PCL_PREDEFINI') ='STD')  then
            begin
            PGIBox ('Vous ne pouvez pas rajouter un cumul standard dans une rubrique', 'Gestion associée');
            exit;
            end;  }

          T := TOB.Create('CUMULRUBRIQUE', P, -1);

          // DANS LE CAS OU LA SOURCE = GRID_CUMULS
          if Source = GRID_CUMULS then TD.ChangeParent(TOB_CUMULCACHE, -1);

          T.PutValue('PCR_CUMULPAIE', TD.GetValue('PCL_CUMULPAIE'));
          T.PutValue('PCR_LIBELLE', TD.GetValue('PCL_LIBELLE'));
          T.PutValue('PCR_RUBRIQUE', RUBRIQUE);
          T.PutValue('PCR_NATURERUB', NATURE);
          //DEB PT-3
          T.PutValue('PCR_PREDEFINI', TD.GetValue('PCL_PREDEFINI'));
          T.PutValue('PCR_NODOSSIER', TD.GetValue('PCL_NODOSSIER'));
          if (Quoi = 'S') and (TD.GetValue('PCL_PREDEFINI') = 'CEG') then //PT-4
            T.PutValue('PCR_PREDEFINI', 'STD'); //PT-4
          if (Quoi <> 'C') and (Quoi <> 'S') then
          begin
            T.PutValue('PCR_PREDEFINI', 'DOS');
            // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
            T.PutValue('PCR_NODOSSIER', PgRendNoDossier());
          end;
          {
          //  PT-2 : 24/08/2001 : PH : Correction mauvaise affectation champ predefini
        if Quoi = 'C' then
          T.PutValue('PCR_PREDEFINI',TD.GetValue('PCL_PREDEFINI'))
          else T.PutValue('PCR_PREDEFINI',TypeP) ;
          //  FIN PT2   T.PutValue ('PCR_PREDEFINI',TD.GetValue('PCL_PREDEFINI'));
        if (Quoi='C') OR (Quoi='S') then T.PutValue('PCR_NODOSSIER','000000')
          else
          begin
          if (V_PGI_Env <> NIL) then T.PutValue('PCR_NODOSSIER',V_PGI_env.nodossier)
           else T.PutValue('PCR_NODOSSIER','000000');
          end;}
        //FIN PT-3

          if Sender = GRID_GAIN then T.PutValue('PCR_SENS', '+') else T.PutValue('PCR_SENS', '-');
        end
        else
        begin
          if (Quoi = 'C') and (TD.GetValue('PCR_PREDEFINI') = 'CEG') and (V_PGI.Debug = FALSE) then
          begin
            PGIBox('Vous ne pouvez pas changer d''affectation un cumul CEGID dans une rubrique CEGID', 'Gestion associée');
            exit;
          end;
          if (DOS) and (TD.GetValue('PCR_PREDEFINI') = 'STD') then
          begin
            PGIBox('Vous ne pouvez pas changer d''affectation un cumul standard dans une rubrique', 'Gestion associée');
            exit;
          end;

          // ON CHANGE LE PARENT
          TD.ChangeParent(P, -1);

          if Sender = GRID_GAIN then TD.PutValue('PCR_SENS', '+') else TD.PutValue('PCR_SENS', '-');
        end;
      end;
    end;
  end; //
  {DEB PT- 1}
  THGrid(Source).ClearSelected;
  {FIN PT- 1}
  AfficheGrids;
end;

procedure TFCumulGestionAssociee.GRID_CUMULSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbLeft then
  begin
    if Sender = GRID_CUMULS then GRID_CUMULS.BeginDrag(TRUE, 10)
    else if Sender = GRID_GAIN then GRID_GAIN.BeginDrag(TRUE, 10)
    else if Sender = GRID_RETENU then GRID_RETENU.BeginDrag(TRUE, 10);
  end;
end;

procedure TFCumulGestionAssociee.VireDejaUtilise;
var
  T, TT: TOB;
  I: INTEGER;
  SENS: string;
begin
  // BOUCLE POUR ALIMENTER LA TOB DES CUMULS CACHES ET POUR ENLEVER DE LA TOB DES CUMULS
  for I := TOB_CUMULS.Detail.Count - 1 downto 0 do
  begin
    T := TOB_CUMULS.Detail[I];

    TT := TOB_CURUB.FindFirst(['PCR_CUMULPAIE'], [T.GetValue('PCL_CUMULPAIE')], FALSE);
    if TT <> nil then
    begin
      // ON AJOUTE DANS TOB_CUMULCACHE
      T.ChangeParent(TOB_CUMULCACHE, 0);
    end;
  end;

  // ON BOUCLE SUR LA TOB DES CUMULS DE LA RUBRIQUE POUR DISPATCH DANS CUGAIN ET CURETENU
  for I := TOB_CURUB.Detail.Count - 1 downto 0 do
  begin
    T := TOB_CURUB.Detail[I];

    SENS := T.GetValue('PCR_SENS');

    if SENS = '+' then T.Changeparent(TOB_CUGAIN, 0) else T.Changeparent(TOB_CURETENU, 0);
  end;
  TOB_CURUB.Free;
end;

procedure TFCumulGestionAssociee.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
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

procedure TFCumulGestionAssociee.sauve;
begin
  if not Modifie then exit;
  ExecuteSQL('DELETE FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="' + NATURE + '" AND ##PCR_PREDEFINI## PCR_RUBRIQUE="' + RUBRIQUE + '"');
  TOB_CUGAIN.SetAllModifie(TRUE);
  TOB_CURETENU.SetAllModifie(TRUE);
  TOB_CUGAIN.InsertDB(nil, TRUE);
  TOB_CURETENU.InsertDB(nil, TRUE);
  Modifie := FALSE;

end;

procedure TFCumulGestionAssociee.BVALIDERClick(Sender: TObject);
begin
  Sauve;
  ModalResult := mrOk;
end;

procedure TFCumulGestionAssociee.BNONAFFECTE_GAINClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_GAIN, GRID_CUMULS, 0, 0);
end;

procedure TFCumulGestionAssociee.BNONAFFECTE_RETENUClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_RETENU, GRID_CUMULS, 0, 0);
end;

procedure TFCumulGestionAssociee.BGAIN_NONAFFECTEClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_CUMULS, GRID_GAIN, 0, 0);
end;

procedure TFCumulGestionAssociee.BRETENU_NONAFFECTEClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_CUMULS, GRID_RETENU, 0, 0);
end;

procedure TFCumulGestionAssociee.BGAIN_RETENUClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_RETENU, GRID_GAIN, 0, 0);
end;

procedure TFCumulGestionAssociee.BRETENU_GAINClick(Sender: TObject);
begin
  DEPOSE_OBJET(GRID_GAIN, GRID_RETENU, 0, 0);
end;

procedure TFCumulGestionAssociee.GRID_CUMULSDblClick(Sender: TObject);
begin
  AglLanceFiche('PAY', 'CUMUL', '', '', 'ACTION=MODIFICATION');
end;

procedure TFCumulGestionAssociee.GRID_GAINDblClick(Sender: TObject);
begin
  AglLanceFiche('PAY', 'CUMUL', '', GRID_GAIN.Cells[0, GRID_GAIN.Row], 'ACTION=MODIFICATION');
end;

procedure TFCumulGestionAssociee.GRID_RETENUDblClick(Sender: TObject);
begin
  AglLanceFiche('PAY', 'CUMUL', '', GRID_RETENU.Cells[0, GRID_RETENU.Row], 'ACTION=MODIFICATION');
end;

procedure TFCumulGestionAssociee.BHELPClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

end.

