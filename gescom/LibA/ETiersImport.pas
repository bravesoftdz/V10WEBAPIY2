unit ETiersImport;

interface

uses Classes, UTOF, UTOB, ETransUtil,DicoAF,Forms, sysutils, Controls, Dialogs,
{$IFDEF EAGLCLIENT}
{$ELSE}
       dbtables,
{$ENDIF}
{$IFDEF DP}
Annoutils,   DpJurOutils ,
{$ENDIF}
     pgienv, paramsoc, HCtrls, Mul, HMsgBox, HStatus, M3FP, HEnt1;

Type
    TOF_ImportETiers = class(TOF)
    private
      FTobETiers , FTobTiers : TOB;
{$IFDEF DP}
      FCodePer : integer;
      FCodeTiers : string;
{$ENDIF}
      procedure ImportETiers;
{$IFDEF DP}
      procedure CreeUnAnnuaire;
{$ENDIF}
    end;

implementation

procedure TOF_ImportETiers.ImportETiers;
Var i : integer;
    F : TFMul ;
    FQ : TQuery;
    stAuxi : String;
    Tobdet : Tob;
begin
F:=TFMul(Ecran);

if (F.FListe.NbSelected = 0) and (not F.FListe.AllSelected) then
  begin
  PGIBoxAF('Veuillez sélectionner les tiers à intégrer', F.Caption);
  exit;
  end;
if PGIAskAF('Intégrer les tiers sélectionnés ?', F.Caption) = mrNo then Exit;

Try
FTobETiers := TOB.Create('Regroupe_ETiers', nil, -1);

if F.FListe.AllSelected then
  begin
       // on doit tout prendre ... pour mise à jour table tiers
  FQ := PrepareSQL('SELECT * FROM ETIERS', true);
  RecupWhereSQL(F.Q, FQ);
  FQ.Open;
  FTobETiers.LoadDetailDB('ETIERS', '', '', FQ, false, true);
  Ferme(FQ);
  F.FListe.AllSelected := false;
  F.bSelectAll.Down := false;
  end else
  begin
  InitMove(F.FListe.NbSelected,'');
  for i := 0 to F.FListe.NbSelected-1 do
    begin
    F.FListe.GotoLeBookMark(i);
    stAuxi := F.Q.FindField('ETI_AUXILIAIRE').AsString;
    if stAuxi <> '' then
        BEGIN
        TobDet := TOB.Create('ETIERS', FTobETiers, -1);
        TobDet.SelectDB('"'+stAuxi+'"', nil);
        END;
    MoveCur(False);
    end;
  FiniMove;
  F.FListe.ClearSelected;
  end;

FTobTiers := TOB.Create('Regroupe_Tiers', nil, -1);
ETableToTable(FTobETiers, FTobTiers); // FTobTiers.detail = les tiers séléctionnées

FTobETiers.DeleteDB(true);
// Les tiers sont créés
FTobTiers.InsertOrUpdateDB(true);
// Si on a le DP, on cree les lignes d'annuaire associées
{$IFDEF DP}
if (GetParamsoc('SO_AFANNUAIREIMPETIERS') = true) then
if (GetParamsoc('SO_AFLIENDP') = true) then
if (V_PGI_Env.InBaseCommune) then
    begin
    // Sur tous les tiers
    for i := 0 to FTobTiers.Detail.Count-1 do
        begin
        FCodeTiers := FTobTiers.Detail[i].GetValue('T_TIERS');
        Transactions(CreeUnAnnuaire,2) ;
        SynchroniseTiers(False, FCodePer, FCodeTiers);
        end;
    end;
{$ENDIF}

finally
FTobTiers.Free;
FTobTiers:=nil;
FTobETiers.Free;
FTobETiers:=nil;
end;
TFMul(Ecran).BChercheClick(Ecran);
PGIInfo('Tiers(s) intégré(s) avec succès', F.Caption);
end;

{$IFDEF DP}
procedure TOF_ImportETiers.CreeUnAnnuaire;
var
TobAnn:TOB;
begin
TobAnn:=TOB.Create('ANNUAIRE', Nil, -1);
Try
TobAnn.PutValue('ANN_TIERS', FCodeTiers);
FCodePer := CalculNumCle('ANN_CODEPER','ANNUAIRE');
TobAnn.PutValue('ANN_CODEPER', FCodePer);
//TobAnn.PutValue('ANN_FAMPER', 'DIV');  //mcd 22/10/03
//TobAnn.PutValue('ANN_TYPEPER', 'CLI_');  //mcd 22/10/03
TobAnn.PutValue('ANN_DATECREATION', V_PGI.DateEntree);
TobAnn.PutValue('ANN_UTILISATEUR', V_PGI.User);
TobAnn.InsertDB(nil);
finally
TobAnn.Free;
end;
end;
{$ENDIF}


procedure AGLImportETiers(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    theTof : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then theTOF := TFMul(F).LaTOF else exit;
if (theTOF is TOF_ImportETiers) then Transactions(TOF_ImportEtiers(theTOF).ImportETiers, 3)
                                else exit;
end;

initialization
RegisterClasses([TOF_ImportETiers]);
RegisterAGLProc('ImportETiers', True, 0, AGLImportETiers);
end.
