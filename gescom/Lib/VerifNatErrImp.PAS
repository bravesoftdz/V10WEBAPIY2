{***********UNITE*************************************************
Auteur  ...... : JT
Créé le ...... : 29/08/2003
Modifié le ... :   /  /
Description .. : Utilitaire de vérification des natures de documents sans
Suite......... : modèle d'impression
Mots clefs ... : VERIFNATERRIMP;UTILITAIRES
*****************************************************************}
Unit VerifNatErrImp ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     FE_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HTb97, UTOF, UTOB,
     ED_Tools, ParamSoc, UtilULst_TOF, UtilUtilitaires;

procedure EntreeVerifNatErrImp;

Implementation

procedure EntreeVerifNatErrImp;
var Qry : TQuery;
    Cpt, QteNatures : integer;
    Msg, Enreg, Tri, Titre, ChpsExport : string;
    TobNatIncorrect, TobTmp : TOB;
    LstResult : TStringList;
begin
  LstResult := TStringList.create;
  TobNatIncorrect := TOB.Create('Erreur Auxiliaire',nil,-1);
  QteNatures := ExecuteSQL('SELECT COUNT(*) FROM PARPIECE');
  InitMoveProgressForm(nil,TraduireMemoire('Traitement'),TraduireMemoire('Vérification en cours.'),QteNatures,False,True);
  Titre := 'Code Libellé                                 Nom modèle document';
  Qry := OpenSQL('SELECT * FROM PARPIECE',True);
  Qry.First;
  while not Qry.Eof do
  begin
    MoveCurProgressForm(Qry.FindField('GPP_LIBELLE').AsString);
    if (Qry.FindField('GPP_IMPMODELE').AsString <> '') and
       (Qry.FindField('GPP_IMPETAT').AsString = '') then
    begin
      Tri := Qry.FindField('GPP_NATUREPIECEG').AsString ;
      Enreg := Format('%-3s',[Qry.FindField('GPP_NATUREPIECEG').AsString])+' '+
               Format('%-40s',[Qry.FindField('GPP_LIBELLE').AsString])+' '+
               RechDom('GCIMPMODELE',Qry.FindField('GPP_IMPMODELE').AsString,False);
      ChpsExport := '';
      TobTmp := UtilTobCreat(TobNatIncorrect,Tri,Titre,Enreg,ChpsExport);
    end;
    Qry.Next;
  end;
  Ferme(Qry);
  if TobNatIncorrect.Detail.count > 0 then
  begin
    UtilTobSort(TobNatIncorrect);
    Msg := TraduireMemoire('Il y a '+IntToStr(TobNatIncorrect.Detail.count)+' nature(s) sans modèle d''impression');
    LstResult.Add(Msg);
    Msg := '';
    LstResult.Add(Msg);
    Msg := Titre;
    LstResult.Add(Msg);
    for Cpt := 0 to TobNatIncorrect.Detail.count - 1 do
    begin
      Msg := TobNatIncorrect.detail[Cpt].GetValue('U_ENREG');
      LstResult.Add(Msg);
    end;
    FiniMoveProgressForm;
    GCLanceFiche_UtilULst('GC','GCUTILULST','','','4',TobNatIncorrect,LstResult,TraduireMemoire('Vérification des natures sans modèle d''impression d''état'),1);
  end else
    FiniMoveProgressForm;
  if TobNatIncorrect <> nil then FreeAndNil(TobNatIncorrect);
  if LstResult <> nil then FreeAndNil(LstResult);
end;

end.
