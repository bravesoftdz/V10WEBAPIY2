unit UtilReglementAffaire;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,  Fe_Main,
      EdtRdoc,
{$ENDIF}
     SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, UtilSais,LettUtil,
     Classes, HMsgBox, SaisComm, ed_tools,LetBatch, FactArticle,FactCpta,
     Paramsoc,EntGc,
     TiersUtil;

procedure CalcReglementSituations (CodeAffaire : String);
function GetReglementAffaire (CodeAffaire : String) : double;
function GetMontantReglementsPiece (ECR:RMvt; Auxiliaire : string) : double;
Procedure GetInfosReglementsPiece (ECR:RMvt; Auxiliaire : string; Var MtRegl, MtAcompte : double ;var DateEcheance : TdateTime;  var DateReglement : TdateTime);
Procedure GetInfosReglementsAcompte (TOBRDA : Tob; Var MtRegl, MtAcompte : double ;var DateEcheance : TdateTime;  var DateReglement : TdateTime);

implementation

procedure AlimenteReglements (TOBECR, TOBReglements : TOB);
var Sql : String;
    Q : TQuery;
    TOBReglement,TOBR : TOB;
    Journal ,Exercice : string;
    Ind,NumPiece,NumLigne : integer;
begin

  if TOBECR.GetValue('E_LETTRAGE') = '' then exit;
  if TOBECR.GetValue('E_COUVERTUREDEV')= 0 then exit;

  TOBR := TOB.Create ('ECRITURE',TOBreglements,-1);

  // reglement de type autre que LCR soit VIR,ESP,etc...
  Sql := 'SELECT * FROM ECRITURE '
       + 'WHERE E_JOURNAL<>"'+ TOBECR.GetValue('E_JOURNAL')   +'" '
       + 'AND E_AUXILIAIRE="'+ TOBECR.GetValue('E_AUXILIAIRE')+'" '
       + 'AND E_GENERAL="'   + TOBECR.GetValue('E_GENERAL')+'" '
       + 'AND E_LETTRAGE="'  + TOBECR.GetValue('E_LETTRAGE')  +'"';

  Q := OpenSql (Sql,True);

  if Q.eof then
     Begin
     ferme(Q);
     exit;
     end;

  TOBReglement := TOB.Create ('LES REGLEMENTS',nil,-1);
  TOBReglement.LoadDetailDB ('ECRITURE','','',Q,false);

  ferme(Q);

  for ind := 0 to TOBReglement.detail.count - 1 do
      begin
      Journal  :=TOBreglement.detail[ind].GetValue('E_JOURNAL');
      Exercice :=TOBreglement.detail[ind].GetValue('E_EXERCICE');
      Numpiece :=TOBreglement.detail[ind].GetValue('E_NUMEROPIECE');
      NumLigne :=TOBreglement.detail[ind].GetValue('E_NUMLIGNE');
      if TOBReglements.findFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE'],[Journal,Exercice,NumPiece,NumLigne],true) = nil then
         begin
         TOBR := TOB.Create ('ECRITURE', TOBReglements, -1);
         TOBR.Dupliquer (TOBreglement.detail[ind],false,true);
         end;
      end;

  TOBReglement.free;
end;

procedure GetReglementsPiece (TOBreglements : TOB; ECR:RMvt; Auxiliaire : string);
var TOBECR : TOB;
    Q : TQuery;
    SQl : String;
begin

  // recupération de l'ecriture de vente TTC sur auxiliaire (tiers)
  Sql:='SELECT ECRITURE.*, MP_CATEGORIE FROM ECRITURE '
          + 'LEFT JOIN MODEPAIE ON E_MODEPAIE=MP_MODEPAIE '
          + 'LEFT JOIN GENERAUX ON E_CONTREPARTIEGEN=G_GENERAL '
          + 'WHERE E_JOURNAL="'+Ecr.JAL+'" AND E_EXERCICE="'+Ecr.EXO+'" '
          + 'AND E_DATECOMPTABLE="'+UsDateTime(Ecr.DateC)+'" AND E_NUMEROPIECE='+InttoStr(Ecr.Num)+' '
          + 'AND E_QUALIFPIECE="'+Ecr.Simul+'" AND E_NATUREPIECE="'+Ecr.Nature+'" '
          + 'AND E_AUXILIAIRE="' + Auxiliaire+'" ';

  Q := OpenSql (Sql,True);

  if Q.eof then
     Begin
     ferme (Q);
     exit;
     end;

  TOBECR := TOB.Create ('ECRITURE',nil,-1);
  TOBECR.SelectDb ('',Q);

  ferme (Q);
  //
  if TOBECR.GetValue ('MP_CATEGORIE') <> 'LCR' then
     AlimenteReglements (TOBECR, TOBReglements)
  else
     begin
     if (TOBECR.GetValue('E_LETTRAGE') <> '') and (TOBECR.GetValue('E_DATEECHEANCE')<=V_PGI.DateEntree) then
        AlimenteReglements (TOBECR, TOBReglements);
     end;

  TOBECR.Free;

end;

function GetMontantReglementsPiece (ECR:RMvt; Auxiliaire : string) : double;
var TOBReglements : TOB;
    Indice : integer;
begin
  result := 0;
  TOBreglements := TOB.Create ('LES REGLEMENTS',nil,-1);
  GetReglementsPiece (TOBreglements, ECR, Auxiliaire);
  for Indice := 0 to TOBReglements.detail.count -1 do
  begin
      result := result + TOBreglements.detail[indice].GetValue('E_COUVERTUREDEV');
      //Modif du 10/01/2008 par FV pour récupération montant exact réglement
      //result := result + TOBreglements.detail[indice].GetValue('E_COUVERTURE');
  end;
  TOBreglements.free;
end;

//FV1 : 13 /11/2014 - FS#1302 - DURET : Rien dans "Reste dû par Affaire"
Procedure GetInfosReglementsAcompte (TOBRDA : Tob; Var MtRegl, MtAcompte : double ;var DateEcheance : TdateTime;  var DateReglement : TdateTime);
var TOBReglements : TOB;
    Indice        : integer;
    Q             : TQuery;
    StSql         : String;
begin

  MtRegl := 0;
  MtAcompte := 0;
  DateEcheance := Idate1900;
  DateReglement := Idate1900;

  TOBreglements := TOB.Create ('LES REGLEMENTS',nil,-1);

  // recupération de l'enregistrement dans la table acompte
  StSql :='SELECT ACOMPTES.*, MP_CATEGORIE FROM ACOMPTES ';
  StSQl := StSQL + 'LEFT JOIN MODEPAIE ON GAC_MODEPAIE=MP_MODEPAIE ';
  StSQL := StSQL + 'WHERE GAC_NATUREPIECEG="' + TOBRDA.Getstring('GP_NATUREPIECEG') + '" ';
  StSQL := StSQL + '  AND GAC_SOUCHE="' + TOBRDA.GetString('GP_SOUCHE') + '"';
  StSQL := StSQL + '  AND GAC_NUMERO="' + TOBRDA.GetString('GP_NUMERO') + '"';
  StSQL := StSQL + '  AND GAC_INDICEG="' + TOBRDA.GetString('GP_INDICEG') + '"';

  Q := OpenSql (StSql,True);

  if Q.eof then
     Begin
     ferme (Q);
     exit;
     end;

  TOBReglements := TOB.Create ('ECRITURE',nil,-1);
  TOBReglements.LoadDetailDB('', '', '', Q, false);

  ferme (Q);

  for Indice := 0 to TOBReglements.detail.count -1 do
  begin
    //Modif du 10/01/2008 par FV pour récupération montant exact réglement
    If TOBreglements.detail[indice].GetString('GAC_ISREGLEMENT')='X' then
      //Gestion Montant Reglement
      MtRegl := MtRegl + TOBreglements.detail[indice].GetValue('GAC_MONTANTDEV')
    Else
      //Gestion acompte sur reglement
      MtAcompte := MtAcompte + TOBreglements.detail[indice].GetValue('GAC_MONTANTDEV');
    //
    //Gestion derniere date
    if TOBreglements.detail[indice].GetString('GAC_NATUREPIECEG')<>'AC' then
    begin
      DateEcheance  := TOBreglements.detail[Indice].GetDateTime('GAC_DATEECHEANCE');
      DateReglement := TOBreglements.detail[Indice].GetDateTime('GAC_DATEECR');
    end;
  end;

  TOBReglements.Free;

end;

Procedure GetInfosReglementsPiece (ECR:RMvt; Auxiliaire : string; Var MtRegl, MtAcompte : double ;var DateEcheance : TdateTime;  var DateReglement : TdateTime);
var TOBReglements : TOB;
    Indice : integer;
begin

  MtRegl := 0;
  MtAcompte := 0;
  DateEcheance := Idate1900;
  DateReglement := Idate1900;

  TOBreglements := TOB.Create ('LES REGLEMENTS',nil,-1);

  GetReglementsPiece(TOBreglements,ECR,Auxiliaire);

  if TOBReglements.detail.count = 0 then
     Begin
     TOBreglements.free;
     exit;
     end;

  for Indice := 0 to TOBReglements.detail.count -1 do
      begin
      //Gestion Montant Reglement
      //Modif du 10/01/2008 par FV pour récupération montant exact réglement
      //MtRegl := MtRegl + TOBreglements.detail[indice].GetValue('E_COUVERTUREDEV');
      //Gestion acompte sur reglement
      If TOBreglements.detail[indice].GetValue('E_NATUREPIECE')='OC' then
      MtAcompte := MtAcompte + TOBreglements.detail[indice].GetValue('E_COUVERTUREDEV')
    else if TOBreglements.detail[indice].GetValue('E_NATUREPIECE')='RC' then
      MtRegl := MtRegl + TOBreglements.detail[indice].GetValue('E_COUVERTUREDEV');
      //Gestion derniere date
      if TOBreglements.detail[indice].GetValue('E_NATUREPIECE')<>'AC' then
        begin
        DateEcheance := TOBreglements.detail[Indice].GetValue('E_DATEECHEANCE');
        DateReglement := TOBreglements.detail[Indice].GetValue('E_DATECOMPTABLE');
        end;
      end;


  TOBreglements.free;

end;

procedure CalcReglementSituations (CodeAffaire : String);

    procedure InitMontantregle(TOBsituations : TOB);
    var Indice : integer;
    		TOBS : TOB;
    begin
      for Indice := 0 to TOBSituations.detail.count -1 do
      begin
      	TOBS := TOBSituations.detail[Indice];
        TOBS.PutValue('BST_MONTANTREGL',0);
      end;
    end;

    procedure CalculReglementPiece (TOBSituations : TOB);
    var Indice : integer;
    		TOBS : TOB;
        Ecr : RMvt;
        QQ : Tquery;
        Sql :String;
        regle : double;
    begin
      for Indice := 0 to TOBSituations.detail.count -1 do
      begin
      	TOBS := TOBSituations.detail[Indice];
        if (not GetParamSoc ('SO_BTCOMPTAREGL')) or
           (GetInfoParPiece (TOBS.getValue('BST_NATUREPIECE'),'GPP_TYPEECRCPTA')='RIE') then
        begin
        	// Acomptes non comptabilisés
          Sql:='SELECT SUM(GAC_MONTANTDEV) AS TOTALREGLE FROM ACOMPTES WHERE GAC_NATUREPIECEG="'+TOBS.GetValue('BST_NATUREPIECE');
        	Sql := Sql + '" AND GAC_SOUCHE="'+TOBS.GetValue('BST_SOUCHE');
          Sql := Sql + '" AND GAC_NUMERO='+InttoStr(TOBS.GetValue('BST_NUMEROFAC'));
          QQ := OpenSql (Sql,True);
          if not QQ.eof then
          begin
            TOBS.PutValue('BST_MONTANTREGL',QQ.findField('TOTALREGLE').asfloat);
          end;
          ferme(QQ);
        end else
        begin
          if TOBS.GetValue('GP_REFCOMPTABLE') <> '' then
          begin
            Ecr:=DecodeRefGCComptable(TOBS.GetValue('GP_REFCOMPTABLE')) ;
            Regle := GetMontantReglementsPiece (ECR,TOBS.GetValue('T_AUXILIAIRE'));
            if Regle > TOBS.GetValue('BST_MONTANTTTC') then
            begin
              TOBS.PutValue('BST_MONTANTREGL',TOBS.GetValue('BST_MONTANTTTC')); // cas de figure ou l'on lettre un reglement avec +ieurs facture 
            end else
            begin
              TOBS.PutValue('BST_MONTANTREGL',Regle);
            end;
            ferme(QQ);
          end;
        end;
      end;
    end;

var TOBSituations : TOB;
		QQ : TQuery ;
    Sql : String;
begin
	if CodeAffaire = '' then exit;
	TOBSituations := TOB.Create ('LES SITUATIONS',nil,-1);
  Sql := 'SELECT S.*,P.GP_REFCOMPTABLE,P.GP_TIERSFACTURE,T.T_AUXILIAIRE FROM BSITUATIONS S '+
  			 'LEFT JOIN PIECE P ON P.GP_NATUREPIECEG=S.BST_NATUREPIECE AND '+
  			 'P.GP_SOUCHE=S.BST_SOUCHE AND P.GP_NUMERO=S.BST_NUMEROFAC '+
  			 'LEFT JOIN TIERS T ON T.T_NATUREAUXI="CLI" AND '+
  			 'T.T_TIERS=P.GP_TIERSFACTURE '+
         'WHERE S.BST_AFFAIRE="'+CodeAffaire+'" AND S.BST_VIVANTE="X"';
  QQ := OpenSql (Sql,True);
  if not QQ.eof then
  begin

  	TOBSituations.LoadDetailDB ('BSITUATIONS','','',QQ,false);
    InitMontantregle(TOBsituations);
    CalculReglementPiece (TOBSituations);
    TOBSituations.UpdateDB (false);
  end;
  ferme (QQ);
  TOBSituations.free;
end;

function GetReglementAffaire (CodeAffaire : String) : double;
var TOBPieces,TOBS,TOBReglements : TOB;
    QQ: TQuery ;
    Sql : String;
    Indice : integer;
    Ecr : RMvt;
    EcrCpta : string;
    TotalRegle : Double;
begin
	result := 0;
  TOBPieces := TOB.Create ('LES PIECES',nil,-1);
  TOBReglements := TOB.Create ('LES ECRITURES',nil,-1);
  Sql := 'SELECT GP_NATUREPIECEG,GP_SOUCHE, GP_NUMERO,GP_REFCOMPTABLE,GP_TIERSFACTURE,GP_TOTALTTCDEV,T_AUXILIAIRE FROM PIECE '+
  			 'LEFT JOIN TIERS ON T_NATUREAUXI="CLI" AND T_TIERS=GP_TIERSFACTURE '+
  			 'WHERE GP_NATUREPIECEG IN ("FBT","FAC","FBC") AND GP_AFFAIRE="'+CodeAffaire+'"';
  QQ := OpenSql (sql,True);
  if not QQ.eof then
  begin
  	TOBPieces.LoadDetailDB ('PIECE','','',QQ,false);
    //
    //
    for Indice := 0 to TOBPieces.detail.count -1 do
    begin
    	TOBS := TOBPieces.detail[Indice];
//
      EcrCpta := GetInfoParPiece (TOBS.getValue('GP_NATUREPIECEG'),'GPP_TYPEECRCPTA');
      if (not GetParamSoc ('SO_BTCOMPTAREGL')) or (EcrCpta='RIE') then
      begin
        // Acomptes non comptabilisés
        Sql:='SELECT SUM(GAC_MONTANTDEV) AS TOTALREGLE FROM ACOMPTES WHERE GAC_NATUREPIECEG="'+TOBS.GetValue('GP_NATUREPIECEG');
        Sql := Sql + '" AND GAC_SOUCHE="'+ TOBS.GetValue('GP_SOUCHE');
        Sql := Sql + '" AND GAC_NUMERO=' + InttoStr(TOBS.GetValue('GP_NUMERO'));
        Sql := Sql + '" AND GAC_INDICEG='+ InttoStr(TOBS.GetValue('GP_INDICEG'));
        QQ := OpenSql (Sql,True);
        if not QQ.eof then
        begin
          result:= result + QQ.findField('TOTALREGLE').asfloat;
        end;
        ferme(QQ);
      end else
      begin
        if (TOBS.GetValue('GP_REFCOMPTABLE') <> '') then
        begin
          Ecr:=DecodeRefGCComptable(TOBS.GetValue('GP_REFCOMPTABLE')) ;
          TotalRegle := GetMontantReglementsPiece (ECR,TOBS.GetValue('T_AUXILIAIRE')) ;
          If TotalRegle > TOBS.GetValue('GP_TOTALTTCDEV') then TotalRegle := TOBS.GetValue('GP_TOTALTTCDEV'); //dans le cas où le reglement aurait été enregistré pour plusieurs factures
          result := result + TotalRegle ;
        end;
      end;
    end;
  end;
  ferme (QQ);
  TOBPieces.free;
  TOBreglements.free;
end;

end.
