{***********UNITE*************************************************
Auteur  ...... : JTR - Gestion Lien GL-ECR
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Détail des lignes de pièces GC liées à une ligne d'écriture
Mots clefs ... : TOF;ECRVERSLIGNES
*****************************************************************}
unit ECRLIGNES_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  Fe_Main,
  db,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  mul,
{$ELSE EAGLCLIENT}
  eMul,
  MainEAGL,
{$ENDIF EAGLCLIENT}
  uTob,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  HTB97,
  UtilLien,
  Vierge,
  Menus,
  EntGC,
  ParamSoc;

procedure LanceZoomLignesGC(TobL: TOB);

type
  TOF_ECRVERSLIGNES = class(TOF)
  private
    GS: THGRID;
    TobE, TobGS: TOB;
    LI_LigEcr: T_Lien;
    bPiece: TMenuItem;
    procedure bPieceClick(Sender: TObject);

  public
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  end;

implementation

procedure LanceZoomLignesGC(TobL: TOB);
begin
  if not GetParamSocSecur('SO_OKLIENLIGECR', false) then exit;
  if TobL.GetString('E_TYPEMVT') <> 'HT' then exit;
  AGLLanceFiche('YY', 'YYECRVERSLIGNES', '', '', IntToStr(LongInt(TobL)));
end;

procedure TOF_ECRVERSLIGNES.OnArgument(S: string);
var Cpt: integer;
  TobEL: TOB;
  Sql: string;
begin
  inherited;
  LI_LigEcr := T_Lien.Create('YLIGECR');
  TobE := Tob(StrToInt(S));
  TobGS := TOB.Create('LES LIGNES', nil, -1);
  bPiece := TMenuItem(GetControl('mnPiece'));
  bPiece.OnClick := bPieceClick;
  GS := THGrid(GetControl('GS'));
  GS.OnDblClick := bPieceClick;
  { Affectation des infos de l'en tête }
  THEdit(GetControl('C_JAL')).Text := RechDom('CPJOURNAL', TobE.GetString('E_JOURNAL'), False);
  THEdit(GetControl('C_NATPCE')).Text := RechDom('CPNATUREPIECECCMP', TobE.GetString('E_NATUREPIECE'), False);
  THEdit(GetControl('C_DATE')).Text := TobE.GetString('E_DATECOMPTABLE');
  THEdit(GetControl('C_NUMP')).Text := TobE.GetString('E_NUMEROPIECE');
  THEdit(GetControl('C_NUML')).Text := TobE.GetString('E_NUMLIGNE');
  THEdit(GetControl('C_GENERAL')).Text := TobE.GetString('E_GENERAL');
  THEdit(GetControl('C_REF')).Text := TobE.GetString('E_REFINTERNE');
  THEdit(GetControl('C_LIB')).Text := TobE.GetString('E_LIBELLE');
  LI_LigEcr.ChargeTob(nil, TobE);
  { Recherche les lignes correspondantes }
  TobEL := TOB.Create('LES LIENS', nil, -1);
  LI_LigEcr.DupliqueTobLien(TobEL);
  for Cpt := 0 to TobEL.Detail.Count - 1 do
  begin
    Sql := 'SELECT * FROM LIGNE WHERE ' + LI_LigEcr.GetWhereFromTob(TobEL.Detail[Cpt], 'GL_');
    TobGS.LoadDetailFromSQL(Sql, True);
  end;
  if assigned(TobEL) then FreeAndNil(TobEL);
  TobGS.PutGridDetail(GS, False, False, ';GL_NATUREPIECEG;GL_NUMERO;GL_DATEPIECE;GL_CODEARTICLE;GL_LIBELLE;GL_QTEFACT;GL_PUHTDEV;GL_TOTALHTDEV', true);
  THNumEdit(GetControl('C_TOTALHT')).Value := TobGS.Somme('GL_TOTALHTDEV', [''], [''], true); ;
  GS.FColAligns[6] := taRightJustify;
  GS.FColAligns[7] := taRightJustify;
  GS.FColAligns[8] := taRightJustify;
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS);
  if TobGS.Detail.Count = 0 then
    PGIBox('Aucune ligne trouvée.');
end;

procedure TOF_ECRVERSLIGNES.OnClose;
begin
  inherited;
  if assigned(TobGS) then FreeAndNil(TobGS);
  LI_LigEcr.Destroy;
end;

procedure TOF_ECRVERSLIGNES.bPieceClick(Sender: TObject);
var CleDoc: R_CleDoc;
begin
  if ((GS.Row <= 0) or (GS.Row > TobGS.Detail.Count)) then Exit;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := TobGS.Detail[GS.Row - 1].GetString('GL_NATUREPIECEG');
  CleDoc.Souche := TobGS.Detail[GS.Row - 1].GetString('GL_SOUCHE');
  CleDoc.DatePiece := TobGS.Detail[GS.Row - 1].GetDateTime('GL_DATEPIECE');
  CleDoc.NumeroPiece := TobGS.Detail[GS.Row - 1].GetInteger('GL_NUMERO');
  CleDoc.Indice := TobGS.Detail[GS.Row - 1].GetInteger('GL_INDICEG');
  CleDoc.NumLigne := TobGS.Detail[GS.Row - 1].GetInteger('GL_NUMLIGNE');
  CleDoc.NumOrdre := TobGS.Detail[GS.Row - 1].GetInteger('GL_NUMORDRE');
  AglLanceFiche('YY', 'YYECRVERSPIECE', '', '',
    CleDoc.NaturePiece + ';' + CleDoc.Souche + ';' +
    FormatDateTime('ddmmyyyy', CleDoc.DatePiece) + ';' + IntToStr(CleDoc.NumeroPiece) +
    ';' + IntToStr(CleDoc.Indice) + ';');
end;

initialization
  registerclasses([TOF_ECRVERSLIGNES]);
end.
