{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPG_CONSULTPRIM ()
Mots clefs ... : TOF;UTOFPG_CONSULTPRIM
*****************************************************************}
unit UTOFPG_CONSULTPRIM;
{
PT1  : 03/05/2006 PH V_65 prise en compte driver dbMSSQL2005
}

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls, HTB97, HSysMenu,
  HEnt1, Ent1,
  HMsgBox,
  UTOF, Utob;

type
  TOF_PG_CONSULTPRIM = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    LaGrille: THGrid;
    TT: TOB;
    procedure VoirHisto(Sender: TObject; Node: TTreeNode);
    procedure AffichTree(Sender: TObject);
    function RendSqlChamp(Champ, Contenu: string): string;
  end;

implementation

uses P5Def, P5Util;

function TransfSub(DD: string; Deb, Long: integer): string;
begin
  dd := dd + ',' + inttostr(deb) + ',' + inttostr(long);
  case V_PGI.Driver of
    dbSQLANY: ; //A FAIRE        majtable
    dbSQLBASE: ; //A FAIRE
    dbINTRBASE: Result := 'MID(' + dd + ')';
    dbMSSQL, dbMSSQL2005, dbSYBASE: Result := 'substring(' + dd + ')'; // PT1
    dbORACLE7, dbORACLE8: Result := 'Substr(' + dd + ')';
    dbPOL: Result := 'Substr(' + dd + ')';
    dbMSACCESS: Result := 'mid(' + dd + ')';
    dbDB2: ; //A FAIRE
    dbINFORMIX: ; //A FAIRE
  else result := dd;
  end;
end;


procedure TOF_PG_CONSULTPRIM.OnNew;
begin
  inherited;
end;

procedure TOF_PG_CONSULTPRIM.OnDelete;
begin
  inherited;
end;

procedure TOF_PG_CONSULTPRIM.OnUpdate;
begin
  inherited;
end;

procedure TOF_PG_CONSULTPRIM.OnLoad;
begin
  inherited;
end;

procedure TOF_PG_CONSULTPRIM.OnArgument(S: string);
var
  BCherChe: TToolbarButton97;
  Num, i: Integer;
begin
  inherited;
  LaGrille := THGrid(GetControl('LAGRILLE'));
  if LaGrille = nil then exit;
  LaGrille.ColAligns[0] := taCenter;
  LaGrille.ColAligns[1] := taLeftJustify;
  for i := 2 to 5 do LaGrille.ColAligns[i] := taRightJustify;

  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
    VisibiliteChampLibreSal(IntToStr(Num), GetControl('PSA_LIBREPCMB' + IntToStr(Num)), GetControl('TPSA_LIBREPCMB' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  BCherche := TToolbarButton97(GetConTrol('BCherche'));
  if BCherche <> nil then
  begin
    BCherche.OnClick := AffichTree;
    AffichTree(nil);
  end;
end;

procedure TOF_PG_CONSULTPRIM.OnClose;
begin
  inherited;
  if TT <> nil then Tt.free;
end;

procedure TOF_PG_CONSULTPRIM.VoirHisto(Sender: TObject; Node: TTreeNode);
var
  i: integer;
  T, TH: TOB;
  st, S1, S2: string;
  Q: TQuery;
  Sal: string;
  HMTrad: THSystemMenu;
begin
  if LaGrille = nil then exit;
  PaieRazPile(LaGrille);
  i := Node.Index;
  if TT.Detail.count = 0 then exit;
  T := TT.detail[i];
  Sal := T.getvalue('PSA_SALARIE');
  Ecran.Caption := 'Consultation historique du salarié ' + T.getValue('PSA_LIBELLE') + ' ' +
    T.getvalue('PSA_PRENOM');
  UpdateCaption(Ecran);
  S1 := TransfSub('PHB_RUBRIQUE', 5, 1);
  S2 := TransfSub('PHB_RUBRIQUE', 1, 4);
  if (Sal = '') or (Sal = 'Collaborateur') then exit;
  st := 'SELECT PHB_DATEFIN,PHB_LIBELLE,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,PHB_MTREM FROM HISTOBULLETIN ';
  st := st + ' LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = ' + S2 + ' WHERE PHB_SALARIE="' + Sal + '" AND PHB_NATURERUB="AAA" AND PRM_IMPRIMABLE = "X"';
  st := st + ' AND (PHB_MTREM<> 0 OR ' + S1 + '=".")';
  st := st + ' AND PHB_DATEDEBUT >="' + USDateTime(StrToDate(GetControlText('DATEDEBUT'))) + '"';
  st := st + ' AND PHB_DATEFIN <="' + USDateTime(StrToDate(GetControlText('DATEFIN'))) + '"';
  st := st + RendSqlChamp('PRM_THEMEREM', GetControlText('PRM_THEMEREM'));
  st := st + ' ORDER BY PHB_RUBRIQUE,PHB_DATEDEBUT';
  Q := OpenSql(St, true);
  FreeAndNIl (TH);
  TH := TOB.Create('Les lignes collaborateurs', nil, -1);
  TH.LoadDetailDB('HISTOBULLETIN', '', '', Q, FALSE);
  TH.PutGridDetail(LaGrille, FALSE, FALSE, 'PHB_DATEFIN;PHB_LIBELLE;PHB_BASEREM;PHB_TAUXREM;PHB_COEFFREM;PHB_MTREM');
  HMTrad.ResizeGridColumns(LaGrille);
  Ferme(Q);
  FreeAndNIl (TH);
end;

procedure TOF_PG_CONSULTPRIM.AffichTree(Sender: TObject);
var
  TVSAL: TTreeview;
  MyTreeNode1: TTreeNode;
  Q: TQuery;
  i, Num: Integer;
  St, st1, st2, St3, st4: string;
begin
  TVSAL := TTreeview(GetControl('TRVSAL'));
  St := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM DEPORTSAL ' +
    'LEFT JOIN SALARIES ON PSA_SALARIE=PSE_SALARIE WHERE PSE_RESPONSVAR = "' + LeSalarie + '"';
  for Num := 1 to 4 do
  begin
    st1 := RendSqlChamp('PSA_TRAVAILN' + IntToStr(Num), GetControlText('PSA_TRAVAILN' + IntToStr(Num)));
    st2 := RendSqlChamp('PSA_LIBREPCMB' + IntToStr(Num), GetControlText('PSA_LIBREPCMB' + IntToStr(Num)));
  end;
  st3 := RendSqlChamp('PSA_CODESTAT', GetControlText('PSA_CODESTAT'));
  st4 := RendSqlChamp('PSA_ETABLISSEMENT', GetControlText('PSA_ETABLISSEMENT'));
  st := st + st1 + st2 + st3 + st4;
  Q := OpenSql(St, true);
  PaieRazPile(LaGrille);
  TT := TOB.Create('Les collaborateurs', nil, -1);
  TT.LoadDetailDB('SALARIES', '', '', Q, FALSE);
  Ferme(Q);
  if TVSAL <> nil then
  begin
    with TVSAL.Items do
    begin
      Clear; { Supprime des noeuds existants }
      MyTreeNode1 := Add(nil, 'Collaborateur'); { Ajoute un noeud à la racine }
      for i := 0 to TT.detail.count - 1 do
      begin
        st := TT.detail[i].GetValue('PSA_LIBELLE') + ' ' + TT.detail[i].GetValue('PSA_PRENOM');
        AddChild(MyTreeNode1, st);
      end;
    end;
    TVSAL.FullExpand;
    TVSAL.OnChange := Voirhisto;
  end;
end;

function TOF_PG_CONSULTPRIM.RendSqlChamp(Champ, Contenu: string): string;
var
  st, S1: string;
  i: Integer;
begin
  result := '';
  if (Contenu = '') or (Copy(Contenu, 1, 2) = '<<') then exit;
  i := 0;
  st := Contenu;
  result := ' AND (';
  s1 := readtokenst(st);
  while S1 <> '' do
  begin
    if i = 0 then result := result + Champ + '="' + S1 + '"'
    else result := result + ' OR ' + Champ + '="' + S1 + '"';
    i := i + 1;
    s1 := readtokenst(st);
  end;
  result := result + ') ';
end;


initialization
  registerclasses([TOF_PG_CONSULTPRIM]);
end.

