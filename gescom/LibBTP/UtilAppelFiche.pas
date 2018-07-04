unit UtilAppelFiche;

interface

uses  Classes,
      AglInit,
      forms,
      sysutils,
      UentCommun,
      HEnt1,
      {$IFNDEF EAGLCLIENT}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      Fe_Main,
      {$ENDIF}
      HCtrls,hmsgbox,
      DialogEx,
      uTOB;

  procedure AppelAnalyseDoc(fform : Tform);
  function AppelRentabiliteDoc(fform : Tform) : boolean;

implementation

uses UtilTOBPiece,facture,factTOB, Grids,BTPUtil,
		 FactOuvrage,FactVariante,UTofListeInv,FactFormule,
     CalcOLEGenericBTP,
     factArticle,
     FactUtil,
     Paramsoc,
     TiersUtil,
     FactCalc,
     UfactExportXLS;

procedure AppelAnalyseDoc(fform : Tform);
var XX : Tffacture;
    TobParam  : TOB;
begin

  TobParam := Tob.Create('LES Parametres', nil, -1);

  TRY
    XX := TFFacture(fform);

    TobParam.AddChampSupValeur('GSLIG', XX.GS.row -1);
    TobParam.AddChampSupValeur('CODETYPE', 'A');
    TobParam.AddChampSupValeur('TYPE', 'Analyse');

    TheTob                  := TobParam;
    TheTob.data             := XX.TheTOBAffaire;
    XX.TheTOBAffaire.data   := XX.TheTOBTiers;
    XX.TheTOBTiers.data     := XX.LaPieceCourante;
    XX.LaPieceCourante.Data := XX.TheTOBOuvrage;
    XX.TheTobOuvrage.Data   := XX.TheTOBArticles;

    AGLLanceFiche('BTP','BTLANCEANALYSE','','','');

  FINALLY
    Thetob := Nil;
    FreeandNil(TobParam);
  END;


end;

function AppelRentabiliteDoc(fform : Tform) : boolean;
var XX : Tffacture;
    TobParam  : TOB;
begin
  result := false;
  TobParam := Tob.Create('LES Parametres', nil, -1);

  TRY
    XX := TFFacture(fform);

    TobParam.AddChampSupValeur('GSLIG', XX.GS.row -1);
    TobParam.AddChampSupValeur('CODETYPE', 'R');
    TobParam.AddChampSupValeur('TYPE', 'Simulation Rentabilité');
    TobParaM.AddChampSupValeur('DEVTAUX', XX.DEV.Taux);
    TobParam.AddChampSupValeur('OKVALIDE', false);

    TheTob                  := TobParam;
    TheTob.data             := XX.TheTOBAffaire;
    XX.TheTOBAffaire.data   := XX.TheTOBTiers;
    XX.TheTOBTiers.data     := XX.LaPieceCourante;
    XX.LaPieceCourante.Data := XX.TheTOBOuvrage;
    XX.TheTobOuvrage.Data   := XX.TheTOBArticles;
    XX.TheTOBArticles.Data  := XX.TheTOBPorcs;
    XX.TheTobPorcs.Data     := XX.ThePieceTrait;
    XX.ThePieceTrait.Data   := XX.TheTOBBases;
    XX.TheTOBBases.Data     := XX.TheTOBBasesL;

    AGLLanceFiche('BTP','BTLANCEANALYSE','','','');
    Result := TobParam.GetBoolean('OKVALIDE')
  FINALLY
    Thetob := Nil;
    FreeandNil(TobParam);
  END;


end;

end.
