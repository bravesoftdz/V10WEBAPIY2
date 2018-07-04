{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM gestion des horaires de l'établissement
Mots clefs ... : PAIE
*****************************************************************}
{
PT1 04/05/2004 V_50 SB FQ 10604 Zone etablissement inaccessible
}
unit UTOMHORAIREETAB;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, DBGrids, Fiche, FichList,
{$ELSE}
  eFiche, eFichList,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97, PgOutils;

type
  TOM_HORAIREETAB = class(TOM)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
  private
    ETAB: string;
  end;

implementation
{ TOM_HORAIREETAB }

procedure TOM_HORAIREETAB.OnArgument(stArgument: string);
var
{$IFNDEF EAGLCLIENT}
  VC: THDBValComboBox;
{$ELSE}
  VC: THValComboBox;
{$ENDIF}

begin
  ETAB := Trim(StArgument);
  ETAB := ReadTokenSt(ETAB); // Recup Etablissement sur lequel on travaille
{$IFNDEF EAGLCLIENT}
  VC := THDBValComboBox(GetControl('PHE_ETABLISSEMENT'));
{$ELSE}
  VC := THValComboBox(GetControl('PHE_ETABLISSEMENT'));
{$ENDIF}
  if VC <> nil then
  begin
    //   VC.ItemIndex:=0 ;
    SetControlEnabled('PHE_ETABLISSEMENT', FALSE);
    if Ecran is TFFicheListe then
    begin
      TFFicheListe(Ecran).FRange := ETAB;
{$IFNDEF EAGLCLIENT}
      TFFicheListe(Ecran).FListe.Columns[0].Visible := FALSE;
{$ELSE}
      TFFicheListe(Ecran).FListe.ColWidths[0] := 0;
{$ENDIF}
    end;
  end;

end;

procedure TOM_HORAIREETAB.OnLoadRecord;
var
{$IFNDEF EAGLCLIENT}
  VC: THDBValComboBox;
{$ELSE}
  VC: THValComboBox;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  VC := THDBValComboBox(GetControl('PHE_ETABLISSEMENT'));
{$ELSE}
  VC := THValComboBox(GetControl('PHE_ETABLISSEMENT'));
{$ENDIF}
  if (VC <> nil) and (DS.State in [dsInsert]) then
  begin
    SetField('PHE_ETABLISSEMENT', ETAB);
  end;
  SetControlEnabled('PHE_ETABLISSEMENT', False); { PT1 }
end;

procedure TOM_HORAIREETAB.OnUpdateRecord;
var
  QQ: TQuery;
  iMax: Integer;
begin
  if DS.State in [dsInsert] then
  begin
    QQ := OpenSQL('SELECT MAX(PHE_ORDREHOR) FROM HORAIREETAB WHERE PHE_ETABLISSEMENT="' + GetField('PHE_ETABLISSEMENT') + '"', TRUE);
    if not QQ.EOF then imax := QQ.Fields[0].AsInteger + 1 else iMax := 1;
    Ferme(QQ);
    SetField('PHE_ORDREHOR', IMax);
    SetField('PHE_ETABLISSEMENT', ETAB);
  end;
end;

initialization
  registerclasses([TOM_HORAIREETAB]);
end.

