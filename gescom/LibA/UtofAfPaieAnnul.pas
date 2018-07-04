{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPAIEANNUL ()
Mots clefs ... : TOF;AFPAIEANNUL
*****************************************************************}
unit UtofAfPaieAnnul;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  dbtables,
  mul,
  FE_Main,
  {$ELSE}
  eMul,
  uTob,
  Maineagl,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF, HQry, HTB97, HStatus,
  ParamSoc, UTofAfBaseCodeAffaire, DicoAf, ConfidentAffaire, AfUtilArticle;

type
  TOF_AFPAIEANNUL = class(TOF_AFBASECODEAFFAIRE)
    procedure OnArgument(S: string); override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  private
    StWhereCritere: string;
    procedure BOuvrirOnClick(Sender: TObject);
    procedure SetAnnul;
    procedure SetAllAnnul;
  end;

procedure AFLanceFiche_AnnulGenerPaie;

const
  TexteMessage: array[1..4] of string = (
    {1}'Aucune ligne d''activité sélectionnée.'
    {2},'Confirmez-vous l''annulation du fichier paie ?#10#13L''activité pourra à nouveau être générée dans la Paie'
    {3},'Impossible d''annuler les lignes d''activité '
    {4},'Erreur Traitement Annulation '
    );     

implementation

procedure TOF_AFPAIEANNUL.OnArgument(S: string);
begin
  inherited;
  SetControlText('ACT_PAIENUMFIC', GetParamSoc('SO_AFLIENPAIEFIC'));
  SetControlProperty('ACT_TYPEARTICLE', 'Plus', ' AND (CO_CODE="PRE" OR CO_CODE="FRA" OR CO_CODE="CTR")');
  TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick;
end;

procedure TOF_AFPAIEANNUL.BOuvrirOnClick(Sender: TObject);
var i,NLigneOk: Integer;
begin
  if (not TFMul(Ecran).fListe.AllSelected) and (TFMul(Ecran).fListe.nbSelected = 0) then
  begin
    PGIBoxAf(TexteMessage[1], Ecran.Caption);  //Pas de lignes sélectionnées
    Exit;
  end else
    if PGIAskAf(TexteMessage[2], Ecran.Caption) <> mrYes then  //demande Confirmation
    Exit;
  NLigneOk := 0;
  try
    with TFMul(Ecran) do
    begin
      if FListe.AllSelected then
      begin
        StWhereCritere := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
        if Transactions(SetAllAnnul, 3) <> oeOK then
          PGIBoxAf(TexteMessage[3], Caption);
        FListe.AllSelected := false;
        TToolBarButton97(GetControl('bSelectAll')).Down := false;
      end else
      if (not fListe.AllSelected) and (fListe.nbSelected <> 0) then
      begin
        InitMove(FListe.NbSelected, '');
        for i := 0 to FListe.NbSelected - 1 do
        begin
          FListe.GotoLeBookMark(i);
          {$IFDEF EAGLCLIENT}
          Q.TQ.Seek(FListe.Row - 1);
          {$ENDIF}
          if Transactions(SetAnnul, 3) = oeOK then inc(NLigneOk);
          MoveCur(False);
        end;
        if NLigneOk < FListe.NbSelected then
          PGIBoxAf(TexteMessage[3], Caption);
        FListe.ClearSelected;
        FiniMove;
      end;
      ChercheClick;
    end;
  except
    on E: Exception do
    begin
      MessageAlerte(TexteMessage[4]+ '#10#13' + E.message);
    end; // on
  end; // try
end;

procedure TOF_AFPAIEANNUL.SetAnnul;
var StSql: string;
begin
  with TFMul(Ecran) do
    StSql := 'UPDATE ACTIVITE SET ACT_PAIENUMFIC=0' +
      ' WHERE ACT_TYPEACTIVITE="' + Q.FindField('ACT_TYPEACTIVITE').AsString + '"' +
      ' AND ACT_AFFAIRE="' + Q.FindField('ACT_AFFAIRE').AsString + '"' +
      ' AND ACT_NUMLIGNEUNIQUE=' + Q.FindField('ACT_NUMLIGNEUNIQUE').AsString;
  ExecuteSQL(StSql);
end;

procedure TOF_AFPAIEANNUL.SetAllAnnul;
var StSql: string;
begin
  StSql := 'UPDATE ACTIVITE SET ACT_PAIENUMFIC=0 ' + StWhereCritere;
  ExecuteSQL(StSql);
end;

procedure TOF_AFPAIEANNUL.NomsChampsAffaire(var Aff, Aff0, Aff1,
  Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff := THEdit(GetControl('ACT_AFFAIRE'));
  Aff1 := THEdit(GetControl('ACT_AFFAIRE1'));
  Aff2 := THEdit(GetControl('ACT_AFFAIRE2'));
  Aff3 := THEdit(GetControl('ACT_AFFAIRE3'));
  Aff4 := THEdit(GetControl('ACT_AVENANT'));
  Tiers := THEdit(GetControl('ACT_TIERS'));

end;

procedure AFLanceFiche_AnnulGenerPaie;
begin
  AGLLanceFiche ('AFF','AFPAIEANNUL_MUL','','','');
end;

initialization
  registerclasses([TOF_AFPAIEANNUL]);
end.
