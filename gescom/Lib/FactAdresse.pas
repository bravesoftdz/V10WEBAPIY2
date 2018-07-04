unit FactAdresse ;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, UTOB, HCtrls, SysUtils, ParamSoc,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     EntGC, TiersUtil, FactTOB, FactComm, UtilPGI,uEntCommun ,UtilTOBPiece,HMsgBox;

procedure LoadLesAdresses(TOBPiece, TOBAdresses: TOB);
function  GetSetNumAdresse: integer;
procedure ValideLesAdresses(TOBPiece, TOBPiece_O, TOBAdresses: TOB);
procedure AddReferenceAdd(TOBPiece, TOBAdresses: TOB);
function  CreerAdresse(TOBAdr: TOB; TypeAdresse: string = 'PIE'): integer;
procedure AffaireVersAdresses(TOBAffaire, TOBAdresses, TOBPiece: TOB);
procedure TiersVersAdresses(TOBTiers, TOBAdresses, TOBPiece: TOB; CodeFactAff: string = '');

function  GetTobAdresseFromAdresses(TobAdresses: Tob; sTypeAdresse, sTiersLivre: string; iNAdresse: integer; sQuelleAdresse: string): boolean;
procedure GetTobPieceAdresseFromTobAdresses(TobPieceAdresse, TobAdresses: Tob);
procedure GetAdrFromCode(TOBAdr: TOB; CodeTiers: string; ForceTobAdresse: Boolean = False);
procedure GetAdrFromTOB(TOBAdr, TOBTiers: TOB; ForceTobAdresse: Boolean = False);
{$IFDEF BTP}
procedure LoadLesAdressesOrigine(TOBPiece, TOBAdresses: TOB);
{$ENDIF}

implementation

uses
  wCommuns,FactUtil;    

procedure LoadLesAdresses(TOBPiece, TOBAdresses: TOB);
var NumL, NumF: Integer;
  Q: TQuery;
  CleDoc: R_CleDoc;
  TOBAdr: TOB;
begin

  if TOBPiece = nil then Exit;

  // Indicateur de mise à jour du tiers selon la nature de pièce
  //{$IFDEF MODE}
  if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_MAJINFOTIERS') = '-' then Exit;
  //{$ENDIF}     

  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    CleDoc := TOB2CleDoc(TOBPiece);
    TOBAdresses.ClearDetail;
    Q := OpenSQL('SELECT * FROM PIECEADRESSE WHERE ' + WherePiece(CleDoc, ttdPieceAdr, False) + ' ORDER BY GPA_TYPEPIECEADR', True,-1, '', True);
    TOBAdresses.LoadDetailDB('PIECEADRESSE', '', '', Q, False);
    Ferme(Q);
    if TOBAdresses.Detail.Count = 1 then
    begin
      TOBAdr := TOB.Create('PIECEADRESSE', TOBAdresses, -1);
      TOBAdr.Dupliquer(TOBAdresses.Detail[0], False, True);
    end;
  end else
  begin
    NumL := TOBPiece.GetValue('GP_NUMADRESSELIVR');
    Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE=' + IntToStr(NumL), True,-1, '', True);
    TOBAdresses.Detail[0].SelectDB('', Q);
    Ferme(Q);
    NumF := TOBPiece.GetValue('GP_NUMADRESSEFACT');
    if NumF = NumL then
    begin
      TOBAdresses.Detail[1].Dupliquer(TOBAdresses.Detail[0], True, True);
    end else
    begin
      Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE=' + IntToStr(NumF), True,-1, '', True);
      TOBAdresses.Detail[1].SelectDB('', Q);
      Ferme(Q);
    end;
  end;
end;

function CreerSoucheADR: integer;
var TOBS: TOB;
  Q: TQuery;
  NumMax: integer;
begin
  Q := OpenSQL('SELECT MAX(ADR_NUMEROADRESSE) FROM ADRESSES', True,-1, '', True);
  if not Q.EOF then NumMax := Q.Fields[0].AsInteger + 1 else NumMax := 1;
  Ferme(Q);
  TOBS := TOB.Create('SOUCHE', nil, -1);
  TOBS.PutValue('SH_TYPE', 'ADR');
  TOBS.PutValue('SH_SOUCHE', 'ADR');
  TOBS.PutValue('SH_LIBELLE', 'Adresses');
  TOBS.PutValue('SH_ABREGE', 'Adresses');
  TOBS.PutValue('SH_NUMDEPART', NumMax);
  TOBS.InsertDB(nil);
  TOBS.Free;
  Result := NumMax;
end;

function GetSetNumAdresse: integer;
var NewNum, ii, Nb, NumMax, NumLu: integer;
  Okok: boolean;
  Q: TQuery;
begin
  ii := 0;
  repeat
    NewNum := -1;
    inc(ii);
    Q := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="ADR" AND SH_SOUCHE="ADR"', True,-1, '', True);
    if not Q.EOF then NewNum := Q.Fields[0].AsInteger;
    Ferme(Q);
    if NewNum = -1 then NewNum := CreerSoucheADR;
    NumLu := NewNum;
    Q := OpenSQL('SELECT MAX(ADR_NUMEROADRESSE) FROM ADRESSES', True,-1, '', True);
    if not Q.EOF then NumMax := Q.Fields[0].AsInteger else NumMax := 0;
    Ferme(Q);
    if NumMax >= NewNum then NewNum := NumMax + 1;
    Nb := ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART=' + IntToStr(NewNum + 1) + ' WHERE SH_TYPE="ADR" AND SH_SOUCHE="ADR" AND SH_NUMDEPART=' +
      IntToStr(NumLu));
    Okok := (Nb = 1);
  until ((Okok) or (ii > 10));
  Result := NewNum;
end;

procedure DetruitAdresse(Num: integer);
begin
  ExecuteSQL('DELETE FROM ADRESSES WHERE ADR_NUMEROADRESSE=' + IntToStr(Num));
end;

procedure AddReferenceAdd(TOBPiece, TOBAdresses: TOB);
var St: string;
begin
  if (TobAdresses = nil) or (Tobpiece = nil) then Exit;
  if TOBAdresses.Detail.count = 0 then Exit;

  St := TOBPiece.GetValue('GP_NATUREPIECEG') + ';' + TOBPiece.GetValue('GP_SOUCHE') + ';' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ';';
  TOBAdresses.Detail[0].PutValue('ADR_REFCODE', St);
  TOBAdresses.Detail[1].PutValue('ADR_REFCODE', St);
end;

function CreerAdresse(TOBAdr: TOB; TypeAdresse: string = 'PIE'): integer;
var NewNum: integer;
begin
  Result := 0;
  NewNum := GetSetNumAdresse;
  if NewNum > 0 then
  begin
    TOBAdr.PutValue('ADR_NUMEROADRESSE', NewNum);
    TOBAdr.PutValue('ADR_TYPEADRESSE', TypeAdresse);
    if TOBAdr.InsertDB(nil) then Result := NewNum;
  end;
end;

procedure ValideLesAdresses(TOBPiece, TOBPiece_O, TOBAdresses: TOB);
var NumL, NumF, i: integer;
  TOBAdr: TOB;
  okok : boolean;
begin

  if (TOBAdresses = nil) or (TobAdresses.Detail.Count = 0) then Exit;

  if TOBPiece = nil then Exit;

  // Indicateur de mise à jour du tiers selon la nature de pièce
  //{$IFDEF MODE}
  if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_MAJINFOTIERS') = '-' then Exit;
  //{$ENDIF}

  NumL := TOBPiece.GetValue('GP_NUMADRESSELIVR');
  NumF := TOBPiece.GetValue('GP_NUMADRESSEFACT');

  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    if ((NumL = NumF) and (TOBAdresses.Detail.Count >= 2)) then TOBAdresses.Detail[1].Free;

    for i := 0 to TOBAdresses.Detail.Count - 1 do
    begin
      TOBAdr := TOBAdresses.Detail[i];
      TOBAdr.PutValue('GPA_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
      TOBAdr.PutValue('GPA_SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
      TOBAdr.PutValue('GPA_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
      TOBAdr.PutValue('GPA_INDICEG', TOBPiece.GetValue('GP_INDICEG'));
      TOBAdr.PutValue('GPA_NUMLIGNE', 0);
      if i = 0 then
        TOBAdr.PutValue('GPA_TYPEPIECEADR', '001')
      else
        TOBAdr.PutValue('GPA_TYPEPIECEADR', '002');
    end;

    //if not TOBAdresses.InsertDB(nil) then
    TRY
      okok := TOBAdresses.InsertOrUpdateDB(True);
    EXCEPT
      on E: Exception do
      begin
        PgiError('Erreur SQL : ' + E.Message, 'Erreur écriture adresses');
      end;
    END;
    if not okok then
    BEGIN
      V_PGI.IoError := oeUnknown;
    END;
  end
  else
  begin
    AddReferenceAdd(TOBPiece, TOBAdresses);
    // Livraison
    if NumL = -1 then
    begin
      // Nouvelle adresse --> Création enregistrement
      NumL := CreerAdresse(TOBAdresses.Detail[0]);
      TOBPiece.PutValue('GP_NUMADRESSELIVR', NumL);
    end else
    begin
      // Adresse existante --> Update enregistrement
      TOBAdresses.Detail[0].PutValue('ADR_NUMEROADRESSE', NumL);
      if TOBAdresses.Detail[0].IsOneModifie then TOBAdresses.Detail[0].UpdateDB(False);
    end;
    // Facturation
    if NumF = -2 then
    begin
      // Nouvelle adresse <> Livr --> Création enregistrement
      NumF := CreerAdresse(TOBAdresses.Detail[1]);
      TOBPiece.PutValue('GP_NUMADRESSEFACT', NumF);
    end else if NumF = -1 then
    begin
      // Nouvelle adresse et = Livr --> Update entête, pas de création enreg adresse
      TOBPiece.PutValue('GP_NUMADRESSEFACT', NumL);
    end else
    begin
      // Ancienne adresse fact ou rendue identique à ancienne adresse livr
      TOBPiece.PutValue('GP_NUMADRESSEFACT', NumF);
      if NumF = NumL then
      begin
        if TOBPiece_O.GetValue('GP_NUMADRESSELIVR') <> TOBPiece_O.GetValue('GP_NUMADRESSEFACT') then DetruitAdresse(TOBPiece_O.GetValue('GP_NUMADRESSEFACT'));
      end else
      begin
        TOBAdresses.Detail[1].PutValue('ADR_NUMEROADRESSE', NumF);
        if TOBAdresses.Detail[1].IsOneModifie then
        begin
          TRY
          TOBAdresses.Detail[1].UpdateDB(False);
          EXCEPT
            on E: Exception do
            begin
              PgiError('Erreur SQL : ' + E.Message, 'Ecritures adresses');
              V_PGI.IoError := oeUnknown;
        end;
          END;
      end;
    end;
  end;
  end;

end;

procedure GetAdrFromTOB(TOBAdr, TOBTiers: TOB; ForceTobAdresse: Boolean = False); 

  procedure PutInTobAdresse; 
  begin
    TOBAdr.PutValue('ADR_LIBELLE',    TOBTiers.GetValue('T_LIBELLE'));
    TOBAdr.PutValue('ADR_LIBELLE2',   TOBTiers.GetValue('T_PRENOM'));
    TOBAdr.PutValue('ADR_JURIDIQUE',  TOBTiers.GetValue('T_JURIDIQUE'));
    TOBAdr.PutValue('ADR_ADRESSE1',   TOBTiers.GetValue('T_ADRESSE1'));
    TOBAdr.PutValue('ADR_ADRESSE2',   TOBTiers.GetValue('T_ADRESSE2'));
    TOBAdr.PutValue('ADR_ADRESSE3',   TOBTiers.GetValue('T_ADRESSE3'));
    TOBAdr.PutValue('ADR_CODEPOSTAL', TOBTiers.GetValue('T_CODEPOSTAL'));
    TOBAdr.PutValue('ADR_VILLE',      TOBTiers.GetValue('T_VILLE'));
    TOBAdr.PutValue('ADR_PAYS',       TOBTiers.GetValue('T_PAYS'));
    TOBAdr.PutValue('ADR_TELEPHONE',  TOBTiers.GetValue('T_TELEPHONE'));
    TobAdr.PutValue('ADR_INCOTERM',   TobTiers.GetValue('YTC_INCOTERM'));
    TobAdr.PutValue('ADR_MODEEXP',    TobTiers.GetValue('YTC_MODEEXP'));
    TobAdr.PutValue('ADR_LIEUDISPO',  TobTiers.GetValue('YTC_LIEUDISPO'));
    TobAdr.PutValue('ADR_EAN',        TobTiers.GetValue('T_EAN'));
    TobAdr.PutValue('ADR_NIF',        TobTiers.GetValue('T_NIF'));
    TobAdr.PutValue('ADR_REGION',     TobTiers.GetValue('T_REGION'));
    //
    TobAdr.PutValue('ADR_NUMEROCONTACT', wGetSqlFieldValue('C_NUMEROCONTACT', 'CONTACT', 'C_AUXILIAIRE = "' + TOBTiers.GetValue('T_AUXILIAIRE') + '" AND C_PRINCIPAL = "' + wTrue + '"'));
  end;

begin

  if ForceTobAdresse then
    PutInTobAdresse
  else
  begin
    if GetParamSoc('SO_GCPIECEADRESSE') then
    begin
      TOBAdr.PutValue('GPA_LIBELLE',    TOBTiers.GetValue('T_LIBELLE'));
      TOBAdr.PutValue('GPA_LIBELLE2',   TOBTiers.GetValue('T_PRENOM'));
      TOBAdr.PutValue('GPA_JURIDIQUE',  TOBTiers.GetValue('T_JURIDIQUE'));
      TOBAdr.PutValue('GPA_ADRESSE1',   TOBTiers.GetValue('T_ADRESSE1'));
      TOBAdr.PutValue('GPA_ADRESSE2',   TOBTiers.GetValue('T_ADRESSE2'));
      TOBAdr.PutValue('GPA_ADRESSE3',   TOBTiers.GetValue('T_ADRESSE3'));
      TOBAdr.PutValue('GPA_CODEPOSTAL', TOBTiers.GetValue('T_CODEPOSTAL'));
      TOBAdr.PutValue('GPA_VILLE',      TOBTiers.GetValue('T_VILLE'));
      TOBAdr.PutValue('GPA_PAYS',       TOBTiers.GetValue('T_PAYS'));
      TobAdr.PutValue('GPA_INCOTERM',   TobTiers.GetValue('YTC_INCOTERM'));
      TobAdr.PutValue('GPA_MODEEXP',    TobTiers.GetValue('YTC_MODEEXP'));
      TobAdr.PutValue('GPA_LIEUDISPO',  TobTiers.GetValue('YTC_LIEUDISPO'));
      TobAdr.PutValue('GPA_EAN',        TobTiers.GetValue('T_EAN'));
      TobAdr.PutValue('GPA_NIF',        TobTiers.GetValue('T_NIF'));
      TobAdr.PutValue('GPA_REGION',     TobTiers.GetValue('T_REGION'));
      //
      TobAdr.PutValue('GPA_NUMEROCONTACT', wGetSqlFieldValue('C_NUMEROCONTACT', 'CONTACT', 'C_AUXILIAIRE = "' + TOBTiers.GetValue('T_AUXILIAIRE') + '" AND C_PRINCIPAL = "' + wTrue + '"'));
    end
    else
      PutInTobAdresse;
  end;
end;

procedure GetAdrFromCode(TOBAdr: TOB; CodeTiers: string; ForceTobAdresse: Boolean = False); 
var TOBT: TOB;
  Q: TQuery;
  Req, chp: string;
begin
  if CodeTiers = '' then Exit;
  TOBT := TOB.Create('TIERS', nil, -1);
  chp := 'T_AUXILIAIRE, T_LIBELLE,T_PRENOM,T_JURIDIQUE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,T_PAYS,T_TELEPHONE,T_EAN,T_NIF,T_REGION'
    + ',YTC_INCOTERM,YTC_MODEEXP,YTC_LIEUDISPO';
  Req := 'SELECT ' + chp + ' FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_TIERS="' + CodeTiers + '"';

  Q := OpenSQL(Req, True,-1, '', True);
  TOBT.SelectDB('', Q);
  Ferme(Q);
  GetAdrFromTOB(TOBAdr, TOBT, ForceTobAdresse); 
  TOBT.Free;
end;


procedure AffaireVersAdresses(TOBAffaire, TOBAdresses, TOBPiece: TOB);
var Q: TQuery;
  TOBAdrAff, TOBAdr: TOB;
begin

  if TobAffaire = nil then Exit;

  //FV : 23/05/2013 - FS#504 - DELABOUDINIERE
  if GetInfoParPiece(TOBPIECE.GetString('GP_NATUREPIECEG'),'GPP_VENTEACHAT')='ACH' then exit;

  Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_TYPEADRESSE="AFA" AND ADR_REFCODE="' +
    TobAffaire.GetValue('AFF_AFFAIRE') + '"', True,-1, '', True);
  TOBAdrAff := TOB.Create('ADRESSES', nil, -1);
  try
    if not Q.EOF then
    begin
      TOBAdrAff.SelectDB('', Q);
      TOBAdr := TOBAdresses.Detail[1];
      if GetParamSoc('SO_GCPIECEADRESSE') then
      begin
        TOBAdr.PutValue('GPA_LIBELLE',    TOBAdrAff.GetValue('ADR_LIBELLE'));
        TOBAdr.PutValue('GPA_LIBELLE2',   TOBAdrAff.GetValue('ADR_LIBELLE2'));
        TOBAdr.PutValue('GPA_JURIDIQUE',  TOBAdrAff.GetValue('ADR_JURIDIQUE'));
        TOBAdr.PutValue('GPA_ADRESSE1',   TOBAdrAff.GetValue('ADR_ADRESSE1'));
        TOBAdr.PutValue('GPA_ADRESSE2',   TOBAdrAff.GetValue('ADR_ADRESSE2'));
        TOBAdr.PutValue('GPA_ADRESSE3',   TOBAdrAff.GetValue('ADR_ADRESSE3'));
        TOBAdr.PutValue('GPA_CODEPOSTAL', TOBAdrAff.GetValue('ADR_CODEPOSTAL'));
        TOBAdr.PutValue('GPA_VILLE',      TOBAdrAff.GetValue('ADR_VILLE'));
        TOBAdr.PutValue('GPA_PAYS',       TOBAdrAff.GetValue('ADR_PAYS'));
          // on considère que si on a une adresse de fact dans la mission, elle est forcement # de 'lautrea dresse (=adresse livraison)
          // à ce jour l'adresse livraison n'est pas géré, mais fait ainsi pour éviter tout pb le jour de la mise en place adr livr
          // Inconvenient, on écrit 2 adresses dans la table.. alors que celle de livraison ne sera pas utilisée
        If TobPiece.getvalue('GP_NUMADRESSELIVR') = TobPiece.getvalue('GP_NUMADRESSEFACT') then
           TOBPiece.PutValue('GP_NUMADRESSEFACT', +2); //mcd 02/10/2003 il faut ds cecas indiquer que les adresse st #
      end else
      begin
        TOBAdr.Dupliquer(TobAdrAff, False, True);
        // reinitialisation de la clé
        TOBAdr.PutValue('ADR_NUMEROADRESSE', 0);
        TOBAdr.PutValue('ADR_REFCODE', '');
        TOBAdr.PutValue('ADR_TYPEADRESSE', '');
        TOBAdr.PutValue('ADR_NUMADRESSEREF', 0);
        TOBPiece.PutValue('GP_NUMADRESSEFACT', -2);
      end;
    end;
  finally
    Ferme(Q);
    TobAdrAff.Free;
  end;
end;

Procedure TiersVersAdresses(TOBTiers, TOBAdresses, TOBPiece: TOB; CodeFactAff: string = '');  
var
  AuxiLivr, CodeLivr: string;
  iNAdresseLivr: integer;
  AuxiFact, CodeFact: string;
  iNAdresseFact: integer;
  TobAdresse: tob;
{  TheTOBTiers : TOB;}
//  QQ : TQuery;
begin
  if TOBTiers = nil then Exit;
  if TOBPiece = nil then Exit;

  // Indicateur de mise à jour du tiers selon la nature de pièce
  //{$IFDEF MODE} 
  if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_MAJINFOTIERS') = '-' then Exit;
  //{$ENDIF}

  if TOBTiers.FieldExists('YTC_NADRESSELIV') then
  begin
    iNadresseLivr := TOBTiers.GetValue('YTC_NADRESSELIV');
    iNadresseFact := TOBTiers.GetValue('YTC_NADRESSEFAC');
  end;

  // if (Tobtiers.GetValue('YTC_TIERSLIVRE') <> '') then   // JS 06102003
  if (NullToVide(Tobtiers.GetValue('YTC_TIERSLIVRE')) <> '') then
  begin
    AuxiLivr := NullToVide(TOBTiers.GetValue('YTC_TIERSLIVRE'));
    CodeLivr := TiersAuxiliaire(AuxiLivr, True);
  end
  else
  begin
    AuxiLivr := NullToVide(TOBTiers.GetValue('T_AUXILIAIRE'));
    CodeLivr := TOBTiers.GetValue('T_TIERS');
  end;

  if CodeLivr = '' then Exit;

  AuxiFact := TOBTiers.GetValue('T_FACTURE');
  if ((AuxiFact = '') or (AuxiFact = AuxiLivr)) then
  begin
    AuxiFact := AuxiLivr;
    CodeFact := CodeLivr;
  end
  else
  begin
    CodeFact := TiersAuxiliaire(AuxiFact, True);
  end;
  {$IFDEF AFFAIRE}
  // specif Affaire , on prends le client a facturer issu de l'affaire
  if (CodeFactAff <> '') then
   begin
   CodeFact := CodeFactAff;
   INadresseFact := 0; // mcd 09/10/03 cas ou clt à facturer vient de l'affaire, il ne faut pasprendre en compte le n° d'adresse du clt d'origine.
                        // si on veut pas la suite,prendre le n° d'adresse à facturer du clt fact de l'affaire, il faudra tout revoir pour avoir
                        // dans la tobtiers,dans ce cas, les info du nouv clt à fact ( appel depuis la fct IncidenceLigenaffaire de facture.pas
   end;
  {$ENDIF}

  // Livraison 
  if (iNAdresseLivr = 0) then
  begin
    //FV1 : 16/12/20014 - FS#1258 - VEODIS : Le changement de tiers dans les contrats n'est plus possible
    if TOBADRESSEs.Detail.count > 0 then
    begin
    if Auxilivr = TOBTiers.GetValue('T_AUXILIAIRE') then
      GetAdrFromTOB(TOBAdresses.Detail[0], TOBTiers)
    else
      GetAdrFromCode(TOBAdresses.Detail[0], CodeLivr);
    // GetAdrFromTOB(TOBAdresses.Detail[0], TOBTiers)
    end;
  end
  else
  begin
    //TobAdresse  est une Tob liée à la table ADRESSES
    //TobAdresses est une Tob liée à la table PIECEADRESSE (ou à la table ADRESSES : ancien fonctionnement)
    TobAdresse := Tob.create('_ADRESSES_', nil, -1);
    try
      GetTobAdresseFromAdresses(TobAdresse, 'TIE', CodeLivr, iNAdresseLivr, 'Livraison');
      if (TobAdresses.detail.count > 0) and (TobAdresse.detail.count > 0) then GetTobPieceAdresseFromTobAdresses(TobAdresses.Detail[0], TobAdresse.Detail[0]);
    finally
      TobAdresse.Free;
    end;
  end;
  if GetParamSoc('SO_GCPIECEADRESSE') then TOBPiece.PutValue('GP_NUMADRESSELIVR', +1)
  else TOBPiece.PutValue('GP_NUMADRESSELIVR', -1);
  // Facturation 
  if (CodeFact = CodeLivr) and (iNAdresseLivr = 0) and (iNAdresseFact = 0) then
  begin
    //FV1 : 16/12/20014 - FS#1258 - VEODIS : Le changement de tiers dans les contrats n'est plus possible
    if TOBADRESSEs.Detail.count > 0 then GetAdrFromTOB(TOBAdresses.Detail[1], TOBTiers);
    if GetParamSoc('SO_GCPIECEADRESSE') then TOBPiece.PutValue('GP_NUMADRESSEFACT', +1)
    else TOBPiece.PutValue('GP_NUMADRESSEFACT', -1);
  end else
  begin
    if (iNAdresseFact <> 0) then
    begin
      //TobAdresse  est une Tob liée à la table ADRESSES
      //TobAdresses est une Tob liée à la table PIECEADRESSE (ou à la table ADRESSES : ancien fonctionnement)
      TobAdresse := Tob.create('_ADRESSES_', nil, -1);
      try
        GetTobAdresseFromAdresses(TobAdresse, 'TIE', CodeFact, iNAdresseFact, 'Facturation');
        if (TobAdresses.detail.count > 0) and (TobAdresse.detail.count > 0) then GetTobPieceAdresseFromTobAdresses(TobAdresses.Detail[1], TobAdresse.Detail[0]);
      finally
        TobAdresse.Free;
      end;
    end
    else
    begin
      //FV1 : 16/12/20014 - FS#1258 - VEODIS : Le changement de tiers dans les contrats n'est plus possible
      if TOBADRESSEs.Detail.count > 0 then
      begin
      if CodeFact = TOBTiers.GetValue('T_AUXILIAIRE') then
        GetAdrFromTOB(TOBAdresses.Detail[1], TOBTiers)
      else
        GetAdrFromCode(TOBAdresses.Detail[1], CodeFact);
    end;
    end;
    if GetParamSoc('SO_GCPIECEADRESSE') then TOBPiece.PutValue('GP_NUMADRESSEFACT', +2)
    else TOBPiece.PutValue('GP_NUMADRESSEFACT', -2);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 29/01/2003
Modifié le ... :   /  /
Description .. : Recherche d'une adresse de livraison ou d'une adresse de facturation depuis
Suite ........ :    * le code tiers livré ou du code tiers facturé
Suite ........ :    * le N° d'adresse
Mots clefs ... : ADRESSE; LIVRAISON; FACTURATION
*****************************************************************}

function GetTobAdresseFromAdresses(TobAdresses: Tob; sTypeAdresse, sTiersLivre: string; iNAdresse: integer; sQuelleAdresse: string): boolean; 
//  sQuelleAdresse   : 'Livraison'     pour une adresse de livraison
//                   : 'Facturation'   pour une adresse de facturation
//                   : 'Règlement'     pour une adresse de règlement

var
  sRequete: string;
  Q: tQuery;
begin
  Result := False;
  if (TobAdresses <> nil) then
  begin
    sRequete := 'SELECT * FROM ADRESSES WHERE (ADR_TYPEADRESSE="' + sTypeAdresse + '") and (ADR_REFCODE="' + sTiersLivre + '") and (ADR_NADRESSE=' + IntToStr(iNAdresse) + ')';
    if (pos('Livraison'  , sQuelleAdresse) > 0) then sRequete := sRequete + ' and (ADR_LIVR="X")';
    if (pos('Facturation', sQuelleAdresse) > 0) then sRequete := sRequete + ' and (ADR_FACT="X")';
    if (pos('Règlement'  , sQuelleAdresse) > 0) then sRequete := sRequete + ' and (ADR_REGL="X")';

    if ExisteSQL(sRequete) then
    begin
      Q := OpenSQL(sRequete, True,-1, '', True);
      try
        TobAdresses.LoadDetailDB('ADRESSES', '', '', Q, True, True);
        Result := (TobAdresses.Detail.Count > 0);
      finally
        Ferme(Q);
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 29/01/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Chargement de la Tob de la table PIECEADRESSE à partir
Suite ........ : d'une Tob de la table ADRESSES
Mots clefs ... : ADRESSE; LIVRAISON
*****************************************************************}

procedure GetTobPieceAdresseFromTobAdresses(TobPieceAdresse, TobAdresses: Tob);
begin
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TobPieceAdresse.PutValue('GPA_LIBELLE', TobAdresses.GetValue('ADR_LIBELLE'));
    TobPieceAdresse.PutValue('GPA_LIBELLE2', TobAdresses.GetValue('ADR_LIBELLE2'));
    TobPieceAdresse.PutValue('GPA_JURIDIQUE', TobAdresses.GetValue('ADR_JURIDIQUE'));
    TobPieceAdresse.PutValue('GPA_ADRESSE1', TobAdresses.GetValue('ADR_ADRESSE1'));
    TobPieceAdresse.PutValue('GPA_ADRESSE2', TobAdresses.GetValue('ADR_ADRESSE2'));
    TobPieceAdresse.PutValue('GPA_ADRESSE3', TobAdresses.GetValue('ADR_ADRESSE3'));
    TobPieceAdresse.PutValue('GPA_CODEPOSTAL', TobAdresses.GetValue('ADR_CODEPOSTAL'));
    TobPieceAdresse.PutValue('GPA_VILLE', TobAdresses.GetValue('ADR_VILLE'));
    TobPieceAdresse.PutValue('GPA_PAYS', TobAdresses.GetValue('ADR_PAYS'));
    TobPieceAdresse.PutValue('GPA_INCOTERM', TobAdresses.GetValue('ADR_INCOTERM'));
    TobPieceAdresse.PutValue('GPA_MODEEXP', TobAdresses.GetValue('ADR_MODEEXP'));
    TobPieceAdresse.PutValue('GPA_LIEUDISPO', TobAdresses.GetValue('ADR_LIEUDISPO'));
    TobPieceAdresse.PutValue('GPA_EAN', TobAdresses.GetValue('ADR_EAN'));
    TobPieceAdresse.PutValue('GPA_NIF', TobAdresses.GetValue('ADR_NIF'));
    TobPieceAdresse.PutValue('GPA_REGION', TobAdresses.GetValue('ADR_REGION')); 
    TobPieceAdresse.PutValue('GPA_NUMEROCONTACT', TobAdresses.GetValue('ADR_NUMEROCONTACT'));
  end
  else
  begin
    TobPieceAdresse.PutValue('ADR_LIBELLE', TobAdresses.GetValue('ADR_LIBELLE'));
    TobPieceAdresse.PutValue('ADR_LIBELLE2', TobAdresses.GetValue('ADR_LIBELLE2'));
    TobPieceAdresse.PutValue('ADR_JURIDIQUE', TobAdresses.GetValue('ADR_JURIDIQUE'));
    TobPieceAdresse.PutValue('ADR_ADRESSE1', TobAdresses.GetValue('ADR_ADRESSE1'));
    TobPieceAdresse.PutValue('ADR_ADRESSE2', TobAdresses.GetValue('ADR_ADRESSE2'));
    TobPieceAdresse.PutValue('ADR_ADRESSE3', TobAdresses.GetValue('ADR_ADRESSE3'));
    TobPieceAdresse.PutValue('ADR_CODEPOSTAL', TobAdresses.GetValue('ADR_CODEPOSTAL'));
    TobPieceAdresse.PutValue('ADR_VILLE', TobAdresses.GetValue('ADR_VILLE'));
    TobPieceAdresse.PutValue('ADR_PAYS', TobAdresses.GetValue('ADR_PAYS'));
    TobPieceAdresse.PutValue('ADR_TELEPHONE', TobAdresses.GetValue('ADR_TELEPHONE'));
    TobPieceAdresse.PutValue('ADR_INCOTERM', TobAdresses.GetValue('ADR_INCOTERM')); 
    TobPieceAdresse.PutValue('ADR_MODEEXP', TobAdresses.GetValue('ADR_MODEEXP'));
    TobPieceAdresse.PutValue('ADR_LIEUDISPO', TobAdresses.GetValue('ADR_LIEUDISPO'));
    TobPieceAdresse.PutValue('ADR_EAN', TobAdresses.GetValue('ADR_EAN'));
    TobPieceAdresse.PutValue('ADR_NIF', TobAdresses.GetValue('ADR_NIF'));
    TobPieceAdresse.PutValue('ADR_REGION', TobAdresses.GetValue('ADR_REGION')); 
    TobPieceAdresse.PutValue('ADR_NUMEROCONTACT', TobAdresses.GetValue('ADR_NUMEROCONTACT'));
    //TobPieceAdresse.PutValue('ADR_CONTACT', TobAdresses.GetValue('ADR_CONTACT'));
    //TobPieceAdresse.PutValue('ADR_CONTACTPRENOM', TobAdresses.GetValue('ADR_CONTACTPRENOM')); 
  end;
end;

{$IFDEF BTP}
procedure LoadLesAdressesOrigine(TOBPiece, TOBAdresses: TOB);
var NumL, NumF: Integer;
  Q: TQuery;
  CleDoc: R_CleDoc;
  TOBAdr: TOB;
begin
  if TOBPiece = nil then Exit;

  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    DecodeRefPiece (TOBPiece.GetValue('PIECEORIGINE'),cledoc);
    TOBAdresses.ClearDetail;
    Q := OpenSQL('SELECT * FROM PIECEADRESSE WHERE ' + WherePiece(CleDoc, ttdPieceAdr, False) + ' ORDER BY GPA_TYPEPIECEADR', True,-1, '', True);
    TOBAdresses.LoadDetailDB('PIECEADRESSE', '', '', Q, False);
    Ferme(Q);
    if TOBAdresses.Detail.Count = 1 then
    begin
      TOBAdr := TOB.Create('PIECEADRESSE', TOBAdresses, -1);
      TOBAdr.Dupliquer(TOBAdresses.Detail[0], False, True);
    end;
  end else
  begin
    NumL := TOBPiece.GetValue('GP_NUMADRESSELIVR');
    Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE=' + IntToStr(NumL), True,-1, '', True);
    TOBAdresses.Detail[0].SelectDB('', Q);
    Ferme(Q);
    NumF := TOBPiece.GetValue('GP_NUMADRESSEFACT');
    if NumF = NumL then
    begin
      TOBAdresses.Detail[1].Dupliquer(TOBAdresses.Detail[0], True, True);
    end else
    begin
      Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE=' + IntToStr(NumF), True,-1, '', True);
      TOBAdresses.Detail[1].SelectDB('', Q);
      Ferme(Q);
    end;
  end;
end;
{$ENDIF}


end.
