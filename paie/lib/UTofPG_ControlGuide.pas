{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 04/07/2001
Modifié le ... :   /  /
Description .. : Unit de controle que toutes les racines de compte utilisées
Suite ........ : dans les préventilations par défaut des rubriques sont bien
Suite ........ : codifiées par rapport au guide des écritures défini dans les
Suite ........ : paramètres généraux
Mots clefs ... : PAIE
*****************************************************************
PT1 : 21/02/2006 : RM V_65 : FQ 12282 Ajout NODOSSIER dans le Where ordre SQL
}
unit UTofPG_ControlGuide;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hqry,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, UTOM, Vierge, AGLInit, EntPaie;

type
  TOF_PG_CONTROLGUIDE = class(TOF)
  private
    TVErreur: TTreeview;
    //PT1 procedure AnalyseLeGuide(LeChamp, LaTable: string; T_Guid: TOB; MyTreeNode1: TTreeNode);
    procedure AnalyseLeGuide(LeChamp, LaTable, LeDossier: string; T_Guid: TOB; MyTreeNode1: TTreeNode); //PT1
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation

procedure TOF_PG_CONTROLGUIDE.OnArgument(Arguments: string);
var
  T_Guid: TOB;
  Q: TQuery;
  St, Quoi: string;
  NoJeu: string;
  MyTreeNode1: TTreeNode;
begin
  inherited;
  TVErreur := TTreeview(GetControl('TVERREUR'));
  if TVErreur = nil then exit;
  NoJeu := VH_Paie.PGModeleEcr; // Recup du numéro de jeu defini dans les paramsoc
  if NoJeu = '' then exit; // Pas de jeu ecriture sortie
  Quoi := ReadTokenst(Arguments); // Recupération du type de table de ventilation que l'on traite
  if Quoi = '' then exit;
  with TVErreur.Items do
    MyTreeNode1 := Add(nil, 'Racines de compte non trouvées');

  T_Guid := TOB.Create('Le Guide', nil, -1);
  //PT1 st := 'SELECT * FROM GUIDEECRPAIE WHERE PGC_JEUECR=' + Nojeu + ''; // DB2
  st := 'SELECT * FROM GUIDEECRPAIE WHERE PGC_JEUECR=' + Nojeu + ' And ##PGC_PREDEFINI## '; //PT1
  Q := OpenSql(St, True);
  T_Guid.LoadDetailDb('GUIDEECRPAIE', '', '', Q, FALSE); // stockage des 20 lignes maxi du guide
  Ferme(Q);
  // Fonctions de controles pour chaque champ racine dans chaque table
  if Quoi = 'R' then
  begin
    //PT1 AnalyseLeGuide('PVS_RACINE1', 'VENTIREMPAIE', T_Guid, MyTreeNode1);
    //PT1 AnalyseLeGuide('PVS_RACINE2', 'VENTIREMPAIE', T_Guid, MyTreeNode1);
    AnalyseLeGuide('PVS_RACINE1', 'VENTIREMPAIE', '##PVS_PREDEFINI##', T_Guid, MyTreeNode1); //PT1
    AnalyseLeGuide('PVS_RACINE2', 'VENTIREMPAIE', '##PVS_PREDEFINI##', T_Guid, MyTreeNode1); //PT1
  end
  else
  begin
    if Quoi = 'C' then
    begin
      //PT1 AnalyseLeGuide('PVT_RACINE1', 'VENTICOTPAIE', T_Guid, MyTreeNode1);
      //PT1 AnalyseLeGuide('PVT_RACINE2', 'VENTICOTPAIE', T_Guid, MyTreeNode1);
      AnalyseLeGuide('PVT_RACINE1', 'VENTICOTPAIE', '##PVT_PREDEFINI##', T_Guid, MyTreeNode1); //PT1
      AnalyseLeGuide('PVT_RACINE2', 'VENTICOTPAIE', '##PVT_PREDEFINI##', T_Guid, MyTreeNode1); //PT1
    end
    else
    begin
      if Quoi = 'O' then
      begin
        //PT1 AnalyseLeGuide('PVO_RACINE1', 'VENTIORGPAIE', T_Guid, MyTreeNode1);
        //PT1 AnalyseLeGuide('PVO_RACINE2', 'VENTIORGPAIE', T_Guid, MyTreeNode1);
        AnalyseLeGuide('PVO_RACINE1', 'VENTIORGPAIE', '##PVO_PREDEFINI##', T_Guid, MyTreeNode1); //PT1
        AnalyseLeGuide('PVO_RACINE2', 'VENTIORGPAIE', '##PVO_PREDEFINI##', T_Guid, MyTreeNode1); //PT1
      end;
    end;
  end;
  FreeAndNIL(T_Guid);
end;


//PT1 procedure TOF_PG_CONTROLGUIDE.AnalyseLeGuide(LeChamp, LaTable: string; T_Guid: TOB; MyTreeNode1: TTreeNode);
procedure TOF_PG_CONTROLGUIDE.AnalyseLeGuide(LeChamp, LaTable, LeDossier: string; T_Guid: TOB; MyTreeNode1: TTreeNode); //PT1
var
  TG: TOB;
  st, Laracine: string;
  Q: TQuery;
  OkTrouve: Boolean;
  LeNbre: Integer;
begin
  OkTrouve := FALSE;
  //PT1 st := 'SELECT DISTINCT ' + LeChamp + ' FROM ' + LaTable + ' WHERE ' + LeChamp + ' <> ""';
  st := 'SELECT DISTINCT ' + LeChamp + ' FROM ' + LaTable + ' WHERE ' + LeChamp + ' <> "" And ' + LeDossier + ''; //PT1
  Q := OpenSql(st, True);
  while not Q.Eof do
  begin
    // la requete nous rend au maximum entre 5 et 15 lignes maxi soit 60 caractères
    LaRacine := Q.FindField(LeChamp).AsString;
    if LaRacine <> '' then
    begin
      TG := T_Guid.FindFirst([''], [''], FALSE);
      while TG <> nil do // Erreur Racine non connue dans le guide
      begin
        OkTrouve := FALSE;
        if Copy(LaRacine, 1, 1) = '4' then LeNbre := VH_Paie.PGLongRacin4
        else LeNbre := VH_Paie.PGLongRacine;
        if Copy(LaRacine, 1, 3) = '421' then LeNbre := VH_Paie.PGLongRacine421;
        if (Copy(LaRacine, 1, LeNbre) = Copy(TG.GetValue('PGC_GENERAL'), 1, LeNbre)) then
        begin
          OkTrouve := TRUE;
          break;
        end;
        TG := T_Guid.FindNext([''], [''], FALSE);
      end;
    end;
    if OkTrouve = FALSE then
      // Racine non trouvée dans le guide, il faut retrouver toutes les rubriques ou organisme concernés par cette racine
    begin
      with TVErreur.Items do
        AddChild(MyTreeNode1, 'Racine du compte ' + Copy(LaRacine, 1, 5) + ' est inconnue ?');
    end;
    Q.Next;
  end;
  Ferme(Q);
end;


initialization
  registerclasses([TOF_PG_CONTROLGUIDE]);
end.

