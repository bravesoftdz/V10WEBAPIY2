{***********UNITE*************************************************
Auteur  ...... : Marie-Noelle GARNIER
Créé le ...... : 24/02/2005
Modifié le ... :   /  /
Description .. : génération automatique d'un chainage et des actions associées
               : suppression automatique d'un chainage qd supp. pièce
Mots clefs ... : chainage
*****************************************************************}
unit UtilChainage;

interface

uses
  UTob,
{$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF !EAGLCLIENT}

  HCtrls
{  ,FactPieceContainer}
  ,sysutils,hmsgbox,
  UTom, FactTob, Controls, HEnt1,EntGC,uEntCommun
  ;
(*
procedure RtGenerationChainage(const PieceContainer: TPieceContainer; stChainage : string);
{ mng 23/06/08 fq011;10051 }
procedure RtTestGenerationChainage(const PieceContainer: TPieceContainer);
*)
Procedure BtGenerationChainage(TobPiece : Tob; StParametre : String);
procedure RTRechChainagePiece(FCleDoc : R_Cledoc; stVenteAchat : String; bDelete : boolean);
procedure SetPieceVisee (Nature,Souche,Numero,Indice: String);
function ExistPieceSAV : boolean;
Const
     RCHTableName = 'ACTIONSCHAINEES';

implementation

//GP_20080201_TS_GC15341 >>>
uses
  wCommuns
  ,EntRT
  ,FactUtil
  ;
//GP_20080201_TS_GC15341 <<<
(*
{ mng 23/06/08 fq011;10051 }
procedure RtTestGenerationChainage(const PieceContainer: TPieceContainer);
var stChainage : string;
begin
    if ( (PieceContainer.Action = taCreat) or (PieceContainer.Action = taModif)
       or ( PieceContainer.DuplicPiece ) or (PieceContainer.TransfoPiece ) ) then
    begin
      stChainage:='';
      if PieceContainer.TCPiece.GetValue('GP_DOMAINE') <> '' then
        stChainage := GetInfoParPieceDomaine(PieceContainer.NewNature, PieceContainer.TCPiece.GetValue('GP_DOMAINE'), 'GDP_CHAINAGE');
      if (stChainage='') then
        stChainage := GetInfoParPiece(PieceContainer.TCPiece.GetValue('GP_NATUREPIECEG'), 'GPP_CHAINAGE');
      if ( stChainage <> '') then
       if (PieceContainer.Action = taModif) and not
          ( ( PieceContainer.DuplicPiece ) or (PieceContainer.TransfoPiece ) ) then
         begin
         if (PieceContainer.TCPiece.GetString('GP_ETATVISA') = 'ATT') then
           RTRechChainagePiece(PieceContainer.CleDoc,PieceContainer.TCPiece.GetString('GP_VENTEACHAT'),false);
         end
       else
         RtGenerationChainage(PieceContainer,stChainage);
    end;
end;

procedure RtGenerationChainage(const PieceContainer: TPieceContainer; stChainage : string);
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
        + ';GENECHAINAGE=' + stChainage
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
*)
procedure BtGenerationChainage(TobPiece : Tob; StParametre : String);
var FCleDoc      : R_CleDoc;
    FTomChainage : TOM;
    FTobChainage : Tob;
    StTiers      : String;
    StLibelle    : String;
    StAffaire    : String;
    StRessource  : String;
    stVenteAchat : String;
    stChainage   : string;
begin

    FTomChainage := CreateTom(RCHTableName, nil, False, True);
    FCleDoc := Tob2CleDoc(TobPiece);
    FTobChainage := Tob.Create(RCHTableName, nil, -1);

    //Chargement des parametres dans leur zone respectives
    StTiers      := Trim(ReadTokenSt(StParametre));  
    StLibelle    := Trim(ReadTokenSt(StParametre));
    StAffaire    := Trim(ReadTokenSt(StParametre));
    StRessource  := Trim(ReadTokenSt(StParametre));
    StVenteAchat := Trim(ReadTokenSt(StParametre));
    StChainage   := Trim(ReadTokenSt(StParametre));

    FTomChainage.Argument('GENETIERS='      + StTiers  //TobPiece.GetString('GP_TIERS')
                        + ';GENELIBELLE='   + StLibelle //TobTiers.GetValue('T_LIBELLE')
                        + ';GENEVENACH='    + StVenteAchat
                        + ';GENECHAINAGE='  + StChainage
                        + ';GENEAFFAIRE='   + StAffaire //TobPiece.GetValue('GP_AFFAIRE')
                        + ';GENERESSOURCE=' + StRessource); //TobPiece.GetValue('GP_RESSOURCE'));

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

procedure RTRechChainagePiece(FCleDoc : R_Cledoc; stVenteAchat : String; bDelete : boolean);
var requete,StProduit,x_where,stMess : string;
    exist1 : boolean;
    TobChainage,TobPieces,TobMere,TobActionsChaine,TobTypeEncours,TobTypActEncours : tob;
    i,j : integer;
    Q: TQuery;
begin
  if stVenteAchat = 'VEN' then StProduit:='GRC' else StProduit:='GRF';
  TobChainage := Tob.Create('CHAINAGEPIECES', nil, -1);
  requete:='SELECT RLC_NUMCHAINAGE FROM CHAINAGEPIECES ';
  x_where:='WHERE RLC_PRODUITPGI="'+StProduit+'"';
  x_where:=x_where+' AND RLC_NATUREPIECEG="'+FCleDoc.NaturePiece+'" AND RLC_NUMERO='+IntToStr(FCleDoc.NumeroPiece);
  x_where:=x_where+' AND RLC_SOUCHE="'+FCleDoc.Souche+'" AND RLC_INDICEG='+IntToStr(FCleDoc.Indice);

  if not TobChainage.LoadDetailDBFromSQL('CHAINAGEPIECES', requete+x_where) then
    begin
    TobChainage.free;
    exit;
    end;

  if bDelete then
    begin
    exist1:=false;
    { si au moins 1 chainage à supprimer (1 seule pièce rattachée : message, sinon on se
      contente de supprimer la ligne pièce du chainage }
    for i:=0 to Pred(TobChainage.detail.count) do
      begin
      TobPieces := Tob.Create('CHAINAGEPIECES', TobChainage.detail[i], -1);
      TobPieces.LoadDetailDBFromSQL('CHAINAGEPIECES',
       'SELECT RLC_NUMCHAINAGE FROM CHAINAGEPIECES WHERE RLC_NUMCHAINAGE='+ IntToStr(TobChainage.detail[i].GetInteger('RLC_NUMCHAINAGE'))+' AND RLC_PRODUITPGI="'+StProduit+'"');
      if TobPieces.detail.count = 1 then exist1:=true;
      end;
    if not exist1 then
      ExecuteSQL('DELETE FROM CHAINAGEPIECES '+x_where)
    else
      begin
      if TobChainage.detail.count > 1 then
         stMess:=TraduireMemoire('Cette pièce est rattachée aux chaînages suivants ')
      else
         stMess:=TraduireMemoire('Cette pièce est rattachée au chaînage suivant ');
      for i:=0 to Pred(TobChainage.detail.count) Do
         stMess:=stMess+IntToStr(TobChainage.detail[i].GetValue('RLC_NUMCHAINAGE'))+' ';

      if TobChainage.detail.count > 1 then
         stMess:=stMess+#13#10+TraduireMemoire('Voulez-vous supprimer ces chaînages ainsi que les actions liées ?')
      else
         stMess:=stMess+#13#10+TraduireMemoire('Voulez-vous supprimer ce chaînage ainsi que les actions liées ?');
      if PgiAsk (stMess,Traduirememoire('Suppression chaînage(s) associé(s)')) = mrYes then
        begin
        for i:=0 to Pred(TobChainage.detail.count) Do
          begin
          wDeleteTable('ACTIONS', 'RAC_NUMCHAINAGE='+ TobChainage.detail[i].GetString('RLC_NUMCHAINAGE')+' AND RAC_PRODUITPGI="'+StProduit+'"');
          ExecuteSQL('DELETE FROM ACTIONSCHAINEES WHERE RCH_NUMERO='+IntToStr(TobChainage.detail[i].GetValue('RLC_NUMCHAINAGE')));
          ExecuteSQL('DELETE FROM CHAINAGEPIECES WHERE RLC_NUMCHAINAGE='+IntToStr(TobChainage.detail[i].GetValue('RLC_NUMCHAINAGE')))
          end;
        end;
      end;
    end
  else
    begin
    TobMere := Tob.Create('les chainages', nil, -1);

    VH_RT.TobTypesAction.Load;
    VH_RT.TobTypesChainage.Load;

    for i:=0 to Pred(TobChainage.detail.count) do
      begin
      TobActionsChaine := Tob.Create('ACTIONSCHAINEES', TobMere, -1);
      Q:=OpenSQL('SELECT * FROM ACTIONSCHAINEES WHERE RCH_NUMERO='+ TobChainage.detail[i].GetString('RLC_NUMCHAINAGE')+' AND RCH_PRODUITPGI="'+StProduit+'"',true);
      TobActionsChaine.SelectDB('', Q);
      TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[TobActionsChaine.GetString('RCH_CHAINAGE'),StProduit],TRUE) ;
      if not TobTypeEncours.GetBoolean('RPG_VISA') then continue;
      { maj statut termine }
      TobActionsChaine.SetBoolean ('RCH_TERMINE',false);
      TobActionsChaine.LoadDetailDBFromSQL('ACTIONS',
       'SELECT * FROM ACTIONS WHERE RAC_NUMCHAINAGE='+ TobChainage.detail[i].GetString('RLC_NUMCHAINAGE')+' AND RAC_PRODUITPGI="'+StProduit+'"');
      for j:=0 to Pred(TobActionsChaine.detail.count) do
        begin
        { si action du modèle trouvée : on prend son état (initial), sinon on met "Prévue" }
        TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TobActionsChaine.detail[j].GetString('RAC_TYPEACTION'),TobActionsChaine.detail[j].GetString('RAC_CHAINAGE'),TobActionsChaine.detail[j].GetInteger('RAC_NUMLIGNE')],TRUE) ;
        if Assigned(TobTypeEncours) and (TobActionsChaine.detail[j].GetInteger('RAC_NUMLIGNE')<> 0 ) then
          TobActionsChaine.detail[j].SetString('RAC_ETATACTION',TobTypActEncours.GetString('RPA_ETATACTION'))
        else
          TobActionsChaine.detail[j].SetString('RAC_ETATACTION','PRE');
        end
      end;
    TobMere.InsertOrUpdateDb();
    end;
  TobChainage.free;
end;

procedure SetPieceVisee (Nature,Souche,Numero,Indice: String);
//GP_20080201_TS_GC15341 >>>
var
  SetPieceVivante : String;
  CleDoc: R_CleDoc;
//GP_20080201_TS_GC15341 <<<
begin
  if (GetInfoParPiece(Nature,'GPP_ACTIONFINI')='VIS') then
    SetPieceVivante:=',GP_VIVANTE="-" ' else SetPieceVivante:='';
  ExecuteSQl('UPDATE PIECE set GP_ETATVISA="VIS"'+
                             ',GP_VISEUR="'+V_PGI.User+
                             '",GP_DATEVISA="'+USDateTime(NowH)+'" '+SetPieceVivante
      +' Where GP_ETATVISA="ATT" AND GP_NATUREPIECEG="'+Nature
      +'" AND GP_SOUCHE="'+Souche
      +'" AND GP_NUMERO='+Numero+' AND GP_INDICEG='+Indice);

  //GP_20080201_TS_GC15341 >>>
  { Action après modification du statut "vivante"/"Visa" d'une pièce}
  CleDoc.NaturePiece := Nature;
  CleDoc.Souche      := Souche;
  CleDoc.NumeroPiece := ValeurI(Numero);
  CleDoc.Indice      := ValeurI(Indice);
//  GPDoActionAfterUpdatePiece(CleDoc)
  //GP_20080201_TS_GC15341 <<<
end;

function ExistPieceSAV : boolean;
var TOBNat : tob;
begin
  result:=true;
  TOBNat:=VH_GC.TOBParPiece.FindFirst(['GPP_PIECESAV'],['X'],False) ;
  if not Assigned(TOBNat) then
    begin
    PGIInfo('Vous ne pouvez pas accéder à ce lien.#13#10 Aucune pièce n''est paramétrée en type SAV','');
    result:=false;
    end;
end;

end.
