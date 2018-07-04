{***********UNITE*************************************************
Auteur  ...... : JT
Créé le ...... : 07/07/2003
Modifié le ... :   /  /
Description .. : Utilitaire de vérification de la cohérence des auxiliaire tiers
Mots clefs ... : VERIFAUXITIERS;UTILITAIRES
*****************************************************************}
Unit VerifAuxiTiers ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     FE_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HTb97, UTOF, UTOB,
     ED_Tools, ParamSoc, UtilULst_TOF, UtilUtilitaires;

procedure EntreeVerifAuxiTiers;

Implementation

procedure EntreeVerifAuxiTiers;
var Qry : TQuery;
    Cpt, QteTiers, AuxiLg : integer;
    AuxiBour, AuxiCorrect, Msg, Enreg, Tri, Titre, ChpsExport : string;
    TobAuxiIncorrect, TobTmp : TOB;
//  LstResult : TListBox;
    LstResult : TStringList;
begin
  if PGIAsk(TraduireMemoire('Veuillez confirmer la recherche des auxiliaires erronés')) <> mrYes then
    exit;
//  LstResult := TListBox.create(nil);
  LstResult := TStringList.create;
  TobAuxiIncorrect := TOB.Create('Erreur Auxiliaire',nil,-1);
  QteTiers := ExecuteSQL('SELECT COUNT(*) FROM TIERS');
  InitMoveProgressForm(nil,TraduireMemoire('Traitement'),TraduireMemoire('Vérification en cours.'),QteTiers,False,True);
  AuxiLg := GetParamSoc('SO_LGCPTEAUX');
  AuxiBour := GetParamSoc('SO_BOURREAUX');
  Titre := 'Nat Auxi trouvé       Auxi calculé      Code tiers        Libellé';
  Qry := OpenSQL('SELECT * FROM TIERS',True);
  Qry.First;
  while not Qry.Eof do
  begin
    MoveCurProgressForm(Qry.FindField('T_AUXILIAIRE').AsString);
    {Calcul auxiliaire comme il devrait l'être}
    if (Qry.FindField('T_NATUREAUXI').AsString = 'CLI') or
       (Qry.FindField('T_NATUREAUXI').AsString = 'PRO') or
       (Qry.FindField('T_NATUREAUXI').AsString = 'CON') then
      AuxiCorrect := GetParamSoc('SO_GCPREFIXEAUXI') + Qry.FindField('T_TIERS').AsString
    else if Qry.FindField('T_NATUREAUXI').AsString = 'FOU' then
      AuxiCorrect := GetParamSoc('SO_GCPREFIXEAUXIFOU') + Qry.FindField('T_TIERS').AsString;
    if (length(AuxiCorrect) < AuxiLg) and (AuxiBour <> '') then
    begin
      while length(AuxiCorrect) < AuxiLg do
        AuxiCorrect := AuxiCorrect + AuxiBour;
    end;
    {Test si auxiliaire non conforme}
    if Qry.FindField('T_AUXILIAIRE').AsString <> AuxiCorrect then
    begin
      Tri := Qry.FindField('T_NATUREAUXI').AsString +';'+Qry.FindField('T_AUXILIAIRE').AsString;
      Enreg := Qry.FindField('T_NATUREAUXI').AsString+' '+Format('%-17s',[Qry.FindField('T_AUXILIAIRE').AsString])+' '+
               Format('%-17s',[AuxiCorrect])+' '+Format('%-17s',[Qry.FindField('T_TIERS').AsString])+' '+
               Qry.FindField('T_LIBELLE').AsString;
      ChpsExport := 'T_NATUREAUXI='+Qry.FindField('T_NATUREAUXI').AsString+';'+
                    'T_AUXILIAIRE='+Qry.FindField('T_AUXILIAIRE').AsString+';'+
                    'AUXICORRECT='+AuxiCorrect+';'+
                    'T_TIERS='+Qry.FindField('T_TIERS').AsString+';'+
                    'T_LIBELLE='+Qry.FindField('T_LIBELLE').AsString;
      TobTmp := UtilTobCreat(TobAuxiIncorrect,Tri,Titre,Enreg,ChpsExport);
    end;
    Qry.Next;
  end;
  Ferme(Qry);

  {Il y a des auxiliaires non conforme}
  if TobAuxiIncorrect.Detail.count > 0 then
  begin
    UtilTobSort(TobAuxiIncorrect);
    Msg := TraduireMemoire('Il y a '+IntToStr(TobAuxiIncorrect.Detail.count)+' code(s) auxiliaire non conforme(s)');
    LstResult.Add(Msg);
    Msg := '';
    LstResult.Add(Msg);
    Msg := Titre;
    LstResult.Add(Msg);
    for Cpt := 0 to TobAuxiIncorrect.Detail.count - 1 do
    begin
      Msg := TobAuxiIncorrect.detail[Cpt].GetValue('U_ENREG');
      LstResult.Add(Msg);
    end;
    FiniMoveProgressForm;
    GCLanceFiche_UtilULst('GC','GCUTILULST','','','0',TobAuxiIncorrect,LstResult,TraduireMemoire('Vérification des auxiliaires non conformes'),0);
  end else
    FiniMoveProgressForm;
  if TobAuxiIncorrect <> nil then FreeAndNil(TobAuxiIncorrect);
  if LstResult <> nil then FreeAndNil(LstResult);
end;

end.
