{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPAIEANNUL ()
Mots clefs ... : TOF;BTPAIEANNUL
*****************************************************************}
unit UtofBTPaieAnnul;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul,
  FE_Main,
  {$ELSE}
  eMul,
  Maineagl,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOB, UTOF, HQry, HTB97, HStatus,
  ParamSoc, UTofAfBaseCodeAffaire, Dicobtp, ConfidentAffaire, AfUtilArticle;

type
  TOF_BTPAIEANNUL = class(TOF_AFBASECODEAFFAIRE)
    procedure OnArgument(S: string); override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  private
    StWhereCritere: string;
    procedure BOuvrirOnClick(Sender: TObject);
    procedure SetAnnul;
    procedure AnnulDansPaie;
  end;

procedure BTLanceFiche_AnnulGenerPaie;

const
  TexteMessage: array[1..4] of string = (
    {1}'Aucune ligne d''activité sélectionnée.'
    {2},'Confirmez-vous l''annulation du fichier paie ?#10#13L''activité pourra à nouveau être générée dans la Paie'
    {3},'Impossible d''annuler les lignes d''activité '
    {4},'Erreur Traitement Annulation '
    );     

implementation

procedure TOF_BTPAIEANNUL.OnArgument(S: string);
begin
	fMulDeTraitement := true;
  inherited;
  FTableName := 'CONSOMMATIONS';
  SetControlText('BCO_PAIENUMFIC', GetParamSoc('SO_AFLIENPAIEFIC'));
//  SetControlProperty('ACT_TYPEARTICLE', 'Plus', ' AND (CO_CODE="PRE" OR CO_CODE="FRA" OR CO_CODE="CTR")');
  TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick;
end;

procedure TOF_BTPAIEANNUL.BOuvrirOnClick(Sender: TObject);
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
        InitMove(Q.RecordCount, '');
        Q.First;
        while Not Q.EOF do
        Begin
          if Transactions(SetAnnul, 3) = oeOK then inc(NLigneOk);
          MoveCur(False);
          Q.Next;
        End;
        if NLigneOk < Q.RecordCount then
          PGIBoxAf(TexteMessage[3], Caption);
        FiniMove;
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

procedure TOF_BTPAIEANNUL.SetAnnul;
var StSql: string;
begin
  StSql := 'UPDATE CONSOMMATIONS SET BCO_PAIENUMFIC=0' +
           ' WHERE BCO_AFFAIRE="' + TFMul(Ecran).Q.FindField('BCO_AFFAIRE').AsString + '"' +
           ' AND BCO_NUMMOUV=' +  IntToStr(TFMul(Ecran).Q.FindField('BCO_NUMMOUV').AsInteger) +
           ' AND BCO_INDICE=' + IntToStr(TFMul(Ecran).Q.FindField('BCO_INDICE').AsInteger);
  ExecuteSQL(StSql);

  AnnulDansPaie;
end;

procedure TOF_BTPAIEANNUL.AnnulDansPaie;
var StSql, CodeSal, CodeRes, Mensuel, ProfilSal, TypeHeure, TypeArt, CodeArt, CodeFa1, CodeFa2, CodeFa3, CodeRub: string;
    dDebMois, dFinMois : TDateTime;
    Q : TQuery ;
    TobD2 : TOB;
    i_rub : integer;
begin
  // Annulation des lignes en attente correspondantes dans la paie

  //récupération date début et fin de mois
  dDebMois := DebutDeMois(TFMul(Ecran).Q.FindField('BCO_DATEMOUV').AsDateTime);
  dFinMois := FinDeMois(TFMul(Ecran).Q.FindField('BCO_DATEMOUV').AsDateTime);

  //récupération du code salarié et du profil
  CodeRes:=TFMul(Ecran).Q.FindField('BCO_RESSOURCE').AsString;
  StSql:='SELECT ARS_SALARIE, ARS_MENSUALISE, PSA_PROFIL FROM RESSOURCE LEFT JOIN SALARIES' +
         ' ON ARS_SALARIE=PSA_SALARIE WHERE ARS_RESSOURCE="'+ CodeRes +'"' ;
  Q:=OpenSQL(StSql,TRUE) ;
  if Not Q.EOF then
  Begin
    CodeSal:=Q.Fields[0].AsString;
    Mensuel:=Q.Fields[1].AsString;
    ProfilSal:=Q.Fields[2].AsString;
  end;
  Ferme(Q) ;
  //récupération des codes famille article
  CodeArt:=TFMul(Ecran).Q.FindField('BCO_ARTICLE').AsString;
  StSql:='SELECT GA_FAMILLENIV1, GA_FAMILLENIV2, GA_FAMILLENIV3, GA_TYPEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+CodeArt+'"' ;
  Q:=OpenSQL(StSql,TRUE) ;
  if Not Q.EOF then
  Begin
    CodeFa1:=Q.Fields[0].AsString;
    CodeFa2:=Q.Fields[1].AsString;
    CodeFa3:=Q.Fields[2].AsString;
    TypeArt:=Q.Fields[3].AsString;
  end;
  Ferme(Q) ;
  //récupération du type d'heure
  TypeHeure:=TFMul(Ecran).Q.FindField('BCO_TYPEHEURE').AsString;

  //récupération du code rubrique
  StSql := 'SELECT ACP_RUBRIQUE,ACP_RESSOURCE,ACP_TYPEHEURE,ACP_TYPEARTICLE,ACP_ARTICLE,ACP_PROFIL,' +
           'ACP_FAMILLENIV1,ACP_FAMILLENIV2,ACP_FAMILLENIV3,ACP_MENSUALISE FROM ACTIVITEPAIE'+
           ' WHERE ACP_RESSOURCE="' + CodeSal + '"' +
           ' OR ACP_TYPEARTICLE="' + TypeArt + '"' +
           ' OR ACP_ARTICLE="' + CodeArt + '"';
  if ProfilSal <> '' then
    StSql := StSql +  ' OR ACP_PROFIL="' + ProfilSal + '"' ;
  if TypeHeure <> '' then
    StSql := StSql + ' OR ACP_TYPEHEURE="' + TypeHeure + '"';
  if CodeFa1 <> '' then
    StSql := StSql + ' OR ACP_FAMILLENIV1="' + CodeFa1 + '"';
  if CodeFa2 <> '' then
    StSql := StSql + ' OR ACP_FAMILLENIV2="' + CodeFa2 + '"';
  if CodeFa3 <> '' then
    StSql := StSql + ' OR ACP_FAMILLENIV3="' + CodeFa3 + '"';
  StSql := StSql + ' ORDER BY ACP_RANG ';

  TobD2:=Tob.create('ACTIVITEPAIE',Nil,-1);
  Q:=OpenSQL(StSql,TRUE) ;
  if not Q.eof Then TobD2.LoadDetailDB('ACTIVITEPAIE','','',Q,False);
  Ferme(Q);
  if (TobD2 <> Nil) and (TobD2.Detail.Count <> 0) then
  begin
    for i_rub := TobD2.Detail.Count - 1 downto 0 do
    begin
      if ((trim(TobD2.Detail[i_rub].GetValue('ACP_ARTICLE')) <> '') and
         (CodeArt <> TobD2.Detail[i_rub].GetValue('ACP_ARTICLE'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
      if ((trim(TobD2.Detail[i_rub].GetValue('ACP_TYPEHEURE')) <> '')
              and (TypeHeure <> TobD2.Detail[i_rub].GetValue('ACP_TYPEHEURE'))) then
//      if ((trim(TobD2.GetValue('ACP_TYPEHEURE')) <> '') and (trim(TobD2.Detail[i_rub].GetValue('ACP_TYPEHEURE')) <> '')
//              and (TobD2.GetValue('ACP_TYPEHEURE') <> TobD2.Detail[i_rub].GetValue('ACP_TYPEHEURE'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
      if ((trim(TobD2.Detail[i_rub].GetValue('ACP_RESSOURCE')) <> '')
         and (CodeRes <> TobD2.Detail[i_rub].GetValue('ACP_RESSOURCE'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;

      if ((trim(TobD2.Detail[i_rub].GetValue('ACP_PROFIL')) <> '')
         and (ProfilSal <> TobD2.Detail[i_rub].GetValue('ACP_PROFIL'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
      if ((trim(TobD2.Detail[i_rub].GetValue('ACP_TYPEARTICLE')) <> '')
         and (TypeArt <> TobD2.Detail[i_rub].GetValue('ACP_TYPEARTICLE'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
      if (Mensuel = 'X') and (TobD2.Detail[i_rub].GetValue('ACP_MENSUALISE') = '-') then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
      if ((trim(CodeFa1) <> '') and (trim(TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV1')) <> '')
        and (CodeFa1 <> TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV1'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
      if ((trim(CodeFa2) <> '') and (trim(TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV2')) <> '')
        and (CodeFa2 <> TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV2'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
      if ((trim(CodeFa3) <> '') and (trim(TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV3')) <> '')
        and (CodeFa1 <> TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV3'))) then
      begin
        TobD2.Detail[i_rub].Free;
        Continue;
      end;
    end;

    if TobD2.Detail.Count > 0 then
    begin
      CodeRub:=TobD2.Detail[0].GetValue('ACP_RUBRIQUE');

      StSql := 'DELETE FROM HISTOSAISRUB' +
               ' WHERE PSD_ORIGINEMVT="MHE"' +
               ' AND PSD_SALARIE="' + CodeSal + '"' +
               ' AND PSD_DATEDEBUT="' + USDateTime(dDebMois) + '"' +
               ' AND PSD_DATEFIN="' + USDateTime(dFinMois) + '"' +
               ' AND PSD_RUBRIQUE="' + CodeRub + '"';
      ExecuteSQL(StSql);
    end;

    // Suppression dans PAIEVENTIL
    StSql := 'DELETE FROM PAIEVENTIL' +
             ' WHERE (PAV_NATURE LIKE "PG%" OR PAV_NATURE LIKE "VS%")' +
             ' AND PAV_COMPTE="' + CodeSal + '"' +
             ' AND PAV_DATEDEBUT="' + USDateTime(dDebMois) + '"' +
             ' AND PAV_DATEFIN="' + USDateTime(dFinMois) + '"';
    ExecuteSQL(StSql);

  end;
  TobD2.free;

end;

procedure TOF_BTPAIEANNUL.NomsChampsAffaire(var Aff, Aff0, Aff1,
  Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff := THEdit(GetControl('BCO_AFFAIRE'));
  Aff1 := THEdit(GetControl('BCO_AFFAIRE1'));
  Aff2 := THEdit(GetControl('BCO_AFFAIRE2'));
  Aff3 := THEdit(GetControl('BCO_AFFAIRE3'));
//  Aff4 := THEdit(GetControl('ACT_AVENANT'));
//  Tiers := THEdit(GetControl('ACT_TIERS'));

end;

procedure BTLanceFiche_AnnulGenerPaie;
begin
  AGLLanceFiche ('BTP','BTPAIEANNUL_MUL','','','');
end;

initialization
  registerclasses([TOF_BTPAIEANNUL]);
end.
