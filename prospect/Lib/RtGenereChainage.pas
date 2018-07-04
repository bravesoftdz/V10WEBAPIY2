{***********UNITE*************************************************
Auteur  ...... : Marie-Noelle GARNIER
Créé le ...... : 24/02/2005
Modifié le ... :   /  /
Description .. : génération automatique d'un chainage et des actions associées
Mots clefs ... :
*****************************************************************}
unit RtGenereChainage;

interface

uses
  UTob
  ,Classes
  ,HCtrls,HEnt1
  ,FactPieceContainer
  {$IFNDEF EAGLCLIENT}
    ,db,dbTables
  {$ENDIF}
  ,sysutils,ParamSoc,UtilPGI,hmsgbox,
  EntGC,UTom, FactTob
  ;
procedure RtGenerationChainage(const PieceContainer: TPieceContainer);

Const
     RCHTableName = 'ACTIONSCHAINEES';

implementation

procedure RtGenerationChainage(const PieceContainer: TPieceContainer);
var
    FCleDoc         : R_CleDoc;
    FTomChainage : TOM;
    FTobChainage : Tob;
begin
    FTomChainage := CreateTom(RCHTableName, nil, False, True);
    FCleDoc := Tob2CleDoc(PieceContainer.TCPiece);
    FTobChainage := Tob.Create(RCHTableName, nil, -1);
    FTomChainage.Argument('GENETIERS='+PieceContainer.TCPiece.GetString('GP_TIERS')
        + ';GENELIBELLE=' + PieceContainer.TCTiers.GetValue('T_LIBELLE')
        + ';GENEVENACH=' + PieceContainer.VenteAchat
        + ';GENECHAINAGE=' + GetInfoParPiece(PieceContainer.TCPiece.GetValue('GP_NATUREPIECEG'), 'GPP_CHAINAGE')
        + ';GENEAFFAIRE=' + PieceContainer.TCPiece.GetValue('GP_AFFAIRE')
        + ';GENERESSOURCE=' + PieceContainer.TCPiece.GetValue('GP_RESSOURCE')
        );        
      if FTomChainage.InitTOB(FTobChainage) then
         if FTomChainage.VerifTOB(FTobChainage) then
            begin
            if FTobChainage.InsertDB(Nil,True) then
               begin
               FTomChainage.AfterVerifTOB(FTobChainage);
               if FTobChainage.GetInteger('RCH_NUMERO') <> 0 then
                  ExecuteSQL('INSERT INTO CHAINAGEPIECES (RLC_NUMCHAINAGE,RLC_NATUREPIECEG,RLC_SOUCHE,RLC_NUMERO,RLC_INDICEG,RLC_PRODUITPGI) '
                  +'values ('+IntToStr(FTobChainage.GetInteger('RCH_NUMERO'))
                  +',"'+FCleDoc.NaturePiece
                  +'","'+FCleDoc.Souche
                  +'",'+IntToStr(FCleDoc.NumeroPiece)
                  +','+IntToStr(FCleDoc.Indice)+',"'+FTobChainage.GetString('RCH_PRODUITPGI')+'")');
               end;
            end;
      FTomChainage.free;
      FTobChainage.free;
end;
end.
