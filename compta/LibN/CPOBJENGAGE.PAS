{***********UNITE*************************************************
Auteur  ...... : Thong Hor LIM, Frédéric BELLARD, Heng Lek CHHOEU
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. : Classe de traitement Engagement.
Mots clefs ... :
*****************************************************************}
unit CPOBJENGAGE;

interface

uses
  {$IFNDEF EAGLCLIENT}
    db,
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  Classes,
  UTob,
  Ent1,
  SaisUtil,
  uLibPieceCompta,
  HEnt1,
  utilPGI,
  hCtrls,
  SysUtils,
  uLibEngage,
  Dialogs,  {showmessage}
  uLibEcriture;  {iif}


Type

TPieceEngagement = class(TPieceCompta)
  private
    FEngage : TOB; {extension des lignes engagements}
    function  ChampEngage( vChamp : string ):boolean; {le champ est l'extension d'engagement}

    procedure MajCleInit; {Initialise Cle engagement et clé init si vide}
    procedure SynchroniseTobEngage   ;

  protected
    class procedure InitialisationClass(Piece: TPieceEngagement; vInfoEcr : TInfoEcriture); {FP 18/04/2006}
    procedure Initialisation ; override; {est appelé par le constructeur}
    Procedure ChargeExtension; override; {est appelé par loadFromSQL}
    procedure CFreeExtension ( vTobEcr : Tob ) ; override;

    function  GetSoldeLigne(vTobEcr: TOB) : double; overload;     {FP 18/04/2006}
  public
    destructor  Destroy ; override ;

    function  GetSoldeLigne(vNumLigne : Integer) : double; overload;
    function  GetMontantRAP(vNumLigne : Integer) : double;

    function  GetValue( vNumLigne: Integer; vChamps: String ): Variant; overload; override;
    procedure PutValue( vNumLigne: Integer; vChamps: String; vValeur: Variant ; vBoForce : boolean = False ); overload; override;
    function  NewRecord    ( vNumLigne : Integer = -1 ) : TOB ; override;
    procedure DeleteRecord ( vNumLigne : Integer ) ; override;
    function  Save: Boolean; override; //FP 18/04/2006
    function  SaveEngage : boolean; {Ecrire dans CENGAGEMENT}
    procedure InverserMontant(Negatif:boolean=true);
    function  IsEngagementEstTermine: boolean; {Thl 24/04/2006}
    function  IsEngagementInitial : boolean; {Vrai=engagement initial, basé sur TOB}
    procedure ViderTobExtension(vNumLigne: Integer=-1); {Vider le tob extension engagement}
    procedure ViderDBExtension; {vider l'extension de la base}

    procedure DupliquerPiece     (PieceSrc: TPieceCompta); override;                                               {FP 02/05/2006}
    procedure DupliquerPieceEng  (PieceSrc: TPieceEngagement; RAZExtension: Boolean);                           {FP 02/05/2006}
    procedure DupliquerLigne     ( vNumLigneDst : Integer ; PieceSrc : TPieceCompta ; vNumLigneSrc : Integer ) ; override;

    //FP 18/04/2006 Pas d'impact sur les engagements? procedure DetruitPieceOrigine ;
    class function  EngagementTTC : boolean; {gestion engagement en mode TTC}
    class function CreerPiece( vInfoEcr : TInfoEcriture = nil ) : TPieceCompta;    {FP 18/04/2006}
    class function CreerPieceEng( vInfoEcr : TInfoEcriture = nil ) : TPieceEngagement; {FP 18/04/2006}
end;


function ExtourneEng(
                      Eng: TPieceEngagement;          {Engagement à extourner}
                      EngExtourne: TPieceEngagement;  {Engagement extourné}
                      EngDate : TDateTime)            {Date de pièce extournée}
                      : boolean;
{ 1 étape
    Eng : prend le statut T avec la date Rappro à EngDate.
        Sauvegarde l'extension Eng
  2 étape
    Crée une pièce extourné avec
    Eng : DateComptable à EngDate
          Exercice selon EngDate
          Inverser les montants de toutes les lignes
          Statut engagement à X
          Sauvegarde la nouvelle pièce engagement avec extension.
  Eng est la pièce extournée
}

function ResteEng(Eng:      TPieceEngagement; {Pièce d'origine}
                  EngRest:  TPieceEngagement  {Piece contenant le reste}
                  ) : boolean;            {Eng reste à engager et écrite}
         {Eng est rendu avec Eng reste à engager}

function RapprocherEngFac(
                   PieceEng,                     {Piece engagement à rapprocher}
                   PieceEngSimul,                {Piece engagement simulé}
                   PieceFact: TPieceEngagement;  {Piece facture à rapprocher}
                   vDate : tDateTime)            {date de rapprochement}
                   :boolean; {rapprochement possible et correcte}

function SupprimerEngagement (Eng : tPieceEngagement):boolean;
         {supprimer un engagement de la base}

function DerapprocherEngagement ( Eng : tPieceEngagement ) :boolean;
         {Ramener à l'engagement initial en
            supprimant les engagements extournés et les restes
            maj les montants rapprochés avec les factures}


Procedure EssaiEngagement;

implementation

uses paramSoc ;

{ TObjTraitEng }


class Function TPieceEngagement.EngagementTTC : boolean;
begin
//  Result := true;//False;
  result := GetParamSocSecur('SO_CPGESTENGAGE', '') <> '2' ;
end;

class function TPieceEngagement.CreerPiece(vInfoEcr: TInfoEcriture): TPieceCompta;
begin
  {b FP 18/04/2006: Provient de TPieceCompta.CreerPiece}
  result := TPieceEngagement.Create( 'V_TOBPIECE', nil, -1 ) ;
  InitialisationClass(result as TPieceEngagement, vInfoEcr);
  {e FP 18/04/2006}
end;

class procedure TPieceEngagement.InitialisationClass(Piece: TPieceEngagement; vInfoEcr : TInfoEcriture);
begin
  {b FP 18/04/2006}

  // Création de l'objet TInfoEcriture
  if Assigned( vInfoEcr )
    then Piece.Info := vInfoEcr
    else Piece.Info := TInfoEcriture.Create ;

  Piece.Initialisation ;
  {e FP 18/04/2006}
end;

class function TPieceEngagement.CreerPieceEng(vInfoEcr: TInfoEcriture): TPieceEngagement;
begin
  {b FP 18/04/2006}
  Result := (CreerPiece(vInfoEcr) as TPieceEngagement);
  {e FP 18/04/2006}
end;

Procedure TPieceEngagement.ChargeExtension;
var lQ       : TQuery ;
    lTobEcr  : TOB ;
    lInIndex : integer ;
begin
  inherited;
//  if Action = taCreat then Exit ;

  ViderTobExtension;

  lQ := OpenSelect('SELECT * FROM CENGAGEMENT '
        + ' WHERE CEN_JOURNAL="'      + GetEntete('E_JOURNAL') + '"'
        + ' AND CEN_EXERCICE="'       + GetEntete('E_EXERCICE') + '"'
        + ' AND CEN_DATECOMPTABLE="'  + UsDateTime( GetEntete('E_DATECOMPTABLE') ) + '"'
        + ' AND CEN_NUMEROPIECE='     + InttoStr( GetEntete('E_NUMEROPIECE') )
        + ' AND CEN_QUALIFPIECE="'    + GetEntete('E_QUALIFPIECE') + '"'
        + ' ORDER BY CEN_JOURNAL, CEN_EXERCICE, CEN_DATECOMPTABLE, '
        + ' CEN_NUMEROPIECE, CEN_NUMLIGNE, CEN_NUMECHE', Dossier ) ;
  while not lQ.EOF do
    begin
    lInIndex := lQ.FindField('CEN_NUMLIGNE').AsInteger ;
    lTobEcr  := GetTob( lInIndex ) ;
    if lTobEcr = nil then break;
    CCreateDBTOBEngage(lTobEcr, FEngage, lQ) ;
    lQ.Next ;
    end ;
  ferme(lQ) ;
  MajCleInit;
end;

procedure TPieceEngagement.Initialisation ;
begin
  inherited;
  FEngage:=Tob.Create('V_ENGAGE',nil,-1);
end;

destructor  TPieceEngagement.Destroy ;
begin
  inherited;
  if Assigned(FEngage) then
    FreeAndNil(FEngage);
end;

procedure TPieceEngagement.ViderTobExtension(vNumLigne: Integer=-1); {Vider le tob extension engagement}
var i : integer;
begin
  if Assigned (FEngage) then
    begin
    FEngage.ClearDetail;
    for i:=0 to Detail.Count-1 do
      CPutTobEngage(Detail[i],nil);
    end
end;

Function  TPieceEngagement.ChampEngage ( vChamp : string ):boolean; {le champ est l'extension d'engagement}
begin
  Result:=Copy(vChamp,1,4)='CEN_';
end;


Function TPieceEngagement.GetSoldeLigne(vNumLigne : Integer) : double;
begin
  {b FP 18/04/2006}
  Result := ARRONDI(GetDouble(vNumLigne, 'E_DEBITDEV') - GetDouble(vNumLigne, 'E_CREDITDEV')
            - GetMontantRAP(vNumLigne),V_PGI.OkDecV);
  {e FP 18/04/2006}
end;

{b FP 18/04/2006}
function TPieceEngagement.GetSoldeLigne(vTobEcr: TOB) : double;
var
  i: Integer;
begin
  Result := 0;
  for i:=0 to Detail.Count-1 do
    begin
    if GetTob(i+1) = vTobEcr then
      begin
      Result := GetSoldeLigne(i+1);
      break;
      end;
    end;
end;
{e FP 18/04/2006}

Function  TPieceEngagement.GetValue ( vNumLigne : Integer ; vChamps : String ) : Variant ;
var lEcr : TOB;
begin
  {b FP 18/04/2006}
  if ChampEngage(vChamps) then
    begin
    lEcr:=GetTOB(vNumLigne);
    if lEcr<>nil then
      Result := CGetValueTobEngage(lEcr, vChamps)
    else
      Result:=#0;
    end
  else
    Result := inherited GetValue(vNumLigne, vChamps);
  {e FP 18/04/2006}
end;

Procedure TPieceEngagement.PutValue( vNumLigne : Integer ; vChamps : String ; vValeur : Variant ; vBoForce : boolean ) ;
var lEcr : TOB;
begin
  lEcr:=GetTob(vNumLigne);
  {b FP 18/04/2006}
  if ChampEngage(vChamps) then
    begin
    if lEcr <> nil then
      begin
      if CGetTobEngage(lEcr) = nil then
        CCreateTobEngage(lEcr, FEngage);
      CPutValueTobEngage(lEcr, vChamps, vValeur);
      end;
    end
  else
    begin
    inherited PutValue(vNumLigne, vChamps, vValeur, vBoForce ) ;
    if vChamps = 'E_JOURNAL' then
      CPutValueTOBEngage(lEcr, 'CEN_JOURNAL', vValeur)
    else if vChamps = 'E_EXERCICE' then
      CPutValueTOBEngage(lEcr, 'CEN_EXERCICE', vValeur)
    else if vChamps = 'E_DATECOMPTABLE' then
      CPutValueTOBEngage(lEcr, 'CEN_DATECOMPTABLE', vValeur)
    else if vChamps = 'E_NUMEROPIECE' then
      CPutValueTOBEngage(lEcr, 'CEN_NUMEROPIECE', vValeur)
    else if vChamps = 'E_QUALIFPIECE' then
      CPutValueTOBEngage(lEcr, 'CEN_QUALIFPIECE', vValeur)
    else if vChamps = 'E_NUMLIGNE' then
      CPutValueTOBEngage(lEcr, 'CEN_NUMLIGNE', vValeur)
    else if vChamps = 'E_NUMECHE' then
      CPutValueTOBEngage(lEcr, 'CEN_NUMECHE', vValeur)
    end;
  {e FP 18/04/2006}
end;


function  TPieceEngagement.NewRecord    ( vNumLigne : Integer = -1 ) : TOB ;
var lEcr:TOB;
begin
  lEcr := inherited NewRecord(vNumLigne);
  CCreateTobEngage (lEcr,FEngage);
  Result:=lEcr;
end;

procedure TPieceEngagement.DeleteRecord ( vNumLigne : Integer ) ;
var lEcr:TOB;
begin
  lEcr := GetTob (vNumLigne);
  if lEcr <> nil then
  begin
    lEcr := CGetTobEngage (lEcr);
    if lEcr <> nil then
      lEcr.Free;
    inherited DeleteRecord (vNumLigne);
  end;
end;


Procedure TPieceEngagement.MajCleInit;
var
  lNumLigne : integer;
  MntTVA    : double;
begin
  MntTVA := 0;
  lNumLigne := 1;
  While not IsOut(lNumLigne) do
  begin
    PutValue(lNumLigne, 'CEN_JOURNAL',       GetString(lNumLigne,'E_JOURNAL'));
    PutValue(lNumLigne, 'CEN_EXERCICE',      GetString(lNumLigne,'E_EXERCICE'));
    PutValue(lNumLigne, 'CEN_NUMEROPIECE',   GetInteger(lNumLigne,'E_NUMEROPIECE'));
    PutValue(lNumLigne, 'CEN_NUMLIGNE',      GetInteger(lNumLigne,'E_NUMLIGNE'));
    PutValue(lNumLigne, 'CEN_QUALIFPIECE',   GetString(lNumLigne,'E_QUALIFPIECE'));
    PutValue(lNumLigne, 'CEN_DATECOMPTABLE', GetDateTime(lNumLigne,'E_DATECOMPTABLE'));
    if GetString(lNumLigne, 'E_QUALIFPIECE') = 'P' then
    begin
      if GetValue(lNumLigne, 'CEN_STATUTENG') = '' then
        PutValue(lNumLigne, 'CEN_STATUTENG', 'E');
      if GetValue(lNumLigne, 'CEN_EJOURNAL') = '' then
        PutValue(lNumLigne, 'CEN_EJOURNAL', GetString (lNumLigne,'E_JOURNAL'));
      if GetValue(lNumLigne, 'CEN_EEXERCICE') = '' then
        PutValue(lNumLigne, 'CEN_EEXERCICE', GetString (lNumLigne,'E_EXERCICE'));
      if GetValue(lNumLigne, 'CEN_ENUMEROPIECE')= 0 then
        PutValue(lNumLigne, 'CEN_ENUMEROPIECE', GetInteger (lNumLigne,'E_NUMEROPIECE'));
      if GetValue(lNumLigne, 'CEN_ENUMLIGNE') = 0 then
        PutValue(lNumLigne, 'CEN_ENUMLIGNE', GetInteger (lNumLigne,'E_NUMLIGNE'));
      if GetValue(lNumLigne, 'CEN_EQUALIFPIECE') = '' then
        PutValue(lNumLigne, 'CEN_EQUALIFPIECE', GetString (lNumLigne,'E_QUALIFPIECE'));
      if GetValue(lNumLigne, 'CEN_EDATECOMPTA') = iDate1900 then
        PutValue(lNumLigne, 'CEN_EDATECOMPTA', GetDateTime (lNumLigne,'E_DATECOMPTABLE'));
      if (GetValue(lNumLigne, 'CEN_MONTANTENGINI') = 0) or (GetValue (lNumLigne, 'CEN_STATUTENG') = 'E') then  {FP 18/04/2006}
        PutValue(lNumLigne, 'CEN_MONTANTENGINI', GetDouble(lNumLigne, 'E_DEBITDEV')- GetDouble(lNumLigne, 'E_CREDITDEV'));
    end
    {b FP 18/04/2006: Recherche le montant de la TVA de la pièce}
    else if (GetString(lNumLigne, 'E_QUALIFPIECE') = 'N') and
            (GetString(lNumLigne, 'E_TYPEMVT') = 'TVA')   and
            (not EngagementTTC)                          and
            (GetValue(lNumLigne, 'CEN_MONTANTRAP') = 0) then
    begin
      MntTVA := MntTVA + GetSoldeLigne(lNumLigne);
    end;
    {e FP 18/04/2006}
    Inc(lNumLigne);
  end;

  {b FP 18/04/2006: Pour la facture, il faut déduire la TVA si gestion des engagements HT est cochée}
  if (MntTVA <> 0) then
  begin
    lNumLigne := 1;
    while not IsOut(lNumLigne) do
    begin
      if (GetString(lNumLigne, 'E_TYPEMVT') = 'TTC') then
        PutValue(lNumLigne, 'CEN_MONTANTRAP', MntTVA)
      else if (GetString(lNumLigne, 'E_TYPEMVT') = 'TVA') then
        PutValue(lNumLigne, 'CEN_MONTANTRAP', GetSoldeLigne(lNumLigne));
      Inc(lNumLigne);
    end;
  end;
  {e FP 18/04/2006}
end;

procedure TPieceEngagement.ViderDBExtension; {vider l'extension de la base}
var Req : string;
begin
  Req := 'CEN_JOURNAL="' + GetEnteteS ('E_JOURNAL')+ '"'+
    ' AND CEN_EXERCICE="'+ GetEnteteS ('E_EXERCICE')+ '"'+
    ' AND CEN_DATECOMPTABLE="'+UsDateTime(GetEnteteDt('E_DATECOMPTABLE'))+'"'+
    ' AND CEN_NUMEROPIECE='+ IntToStr (GetEnteteI('E_NUMEROPIECE'))+
    ' AND CEN_QUALIFPIECE="' + GetEnteteS('E_QUALIFPIECE')+ '"';
  ExecuteSQL ( 'DELETE FROM ' + GetTableDossier(Dossier, 'CENGAGEMENT' ) +
    ' WHERE ' + Req) ;
end;

{b FP 18/04/2006}
function TPieceEngagement.Save: Boolean;
begin
  SynchroniseTobEngage   ;
  Result := inherited Save;
  if Result then
    Result := SaveEngage;
end;
{e FP 18/04/2006}

Function  TPieceEngagement.SaveEngage : boolean; {Ecrire dans CENGAGEMENT}
begin
  ViderDBExtension;
{b FP 02/05/2006: enregistre dans la table uniquement les pièces de type N ou P}
  Result := True;
  if (GetEntete('E_QUALIFPIECE') = 'P') or (GetEntete('E_QUALIFPIECE') = 'N') then
{e FP 02/05/2006}
    begin
    MajCleInit;
    Result := FEngage.InsertDB (nil);
    end;
end;

function TPieceEngagement.IsEngagementEstTermine: boolean;
var
  ligne:   integer;
begin
  Result := True;
  for ligne := Detail.Count downto 1 do begin
    if GetSoldeLigne(ligne) <> 0 then  begin
      Result := False;
      Break;
      end;
    end
end;

function  TPieceEngagement.IsEngagementInitial : boolean; {Vrai=engagement initial, basé sur TOB}
begin
//  Result := false;
  if (GetValue (1,'CEN_EJOURNAL') = #0) or
     (GetValue (1,'CEN_EJOURNAL')= '') then Result := true
  else begin
    Result := (GetEntete ('E_JOURNAL') = GetValue (1, 'CEN_EJOURNAL')) and
              (GetEntete ('E_EXERCICE') = GetValue (1, 'CEN_EEXERCICE')) and
              (GetEntete ('E_DATECOMPTABLE') = GetValue (1, 'CEN_EDATECOMPTA')) and
              (GetEntete ('E_NUMEROPIECE') = GetValue (1, 'CEN_ENUMEROPIECE')) and
              (GetEntete ('E_QUALIFPIECE') = GetValue (1, 'CEN_EQUALIFPIECE'));
  end;
end;

Procedure TPieceEngagement.InverserMontant(Negatif : boolean=true);

  Procedure InverserEcr(vEcr : TOB);

    Procedure InverserAna(vAna : TOB);
    var X      : double;
    begin
      if Negatif then
      begin
        vAna.PutValue('Y_DEBIT',-vAna.GetValue('Y_DEBIT'));
        vAna.PutValue('Y_CREDIT',-vAna.GetValue('Y_CREDIT'));
        vAna.PutValue('Y_DEBITDEV',-vAna.GetValue('Y_DEBITDEV'));
        vAna.PutValue('Y_CREDITDEV',-vAna.GetValue('Y_CREDITDEV'));
        vAna.PutValue('Y_QTE1',-vAna.GetValue('Y_QTE1'));
        vAna.PutValue('Y_QTE2',-vAna.GetValue('Y_QTE2'));
        vAna.PutValue('Y_TOTALQTE1',-vAna.GetValue('Y_TOTALQTE1'));
        vAna.PutValue('Y_TOTALQTE2',-vAna.GetValue('Y_TOTALQTE2'));
        vAna.PutValue('Y_TOTALECRITURE',-vAna.GetValue('Y_TOTALECRITURE'));
        vAna.PutValue('Y_TOTALDEVISE',-vAna.GetValue('Y_TOTALDEVISE'));
      end
      else
      begin
        X := vAna.GetValue('Y_DEBIT');
        vAna.PutValue('Y_DEBIT',vAna.GetValue('Y_CREDIT'));
        vAna.PutValue('Y_CREDIT',X);
        X := vAna.GetValue('Y_DEBITDEV');
        vAna.PutValue('Y_DEBITDEV',vAna.GetValue('Y_CREDITDEV'));
        vAna.PutValue('Y_CREDITDEV',X);
      end;
    end;

  var X      : double;
      lAxe,
      lAna   : TOB;
      i,j    : integer;
  begin {InverserEcr}
    if Negatif then
    begin
      vEcr.PutValue('E_DEBIT',-vEcr.GetValue('E_DEBIT'));
      vEcr.PutValue('E_CREDIT',-vEcr.GetValue('E_CREDIT'));
      vEcr.PutValue('E_DEBITDEV',-vEcr.GetValue('E_DEBITDEV'));
      vEcr.PutValue('E_CREDITDEV',-vEcr.GetValue('E_CREDITDEV'));
      vEcr.PutValue('E_QTE1',-vEcr.GetValue('E_QTE1'));
      vEcr.PutValue('E_QTE2',-vEcr.GetValue('E_QTE2'));
    end
    else
    begin
      X := vEcr.GetValue('E_DEBIT');
      vEcr.PutValue('E_DEBIT',vEcr.GetValue('E_CREDIT'));
      vEcr.PutValue('E_CREDIT',X);
      X := vEcr.GetValue('E_DEBITDEV');
      vEcr.PutValue('E_DEBITDEV',vEcr.GetValue('E_CREDITDEV'));
      vEcr.PutValue('E_CREDITDEV',X);
    end;
    for i := 0 to vEcr.Detail.Count - 1 do
    begin
      lAxe := vEcr.Detail[i];
      for j := 0 to lAxe.Detail.Count - 1 do
      begin
        lAna := lAxe.Detail[j];
        InverserAna (lAna);
      end;
    end;
  end;

var lLigne : integer;
    lEcr   : TOB;
    lNum   : integer;
    lEche  : TOB;
    i      : integer;
begin {InverserMontant}
  for lLigne:=0 to Detail.Count-1 do
  begin
    lEcr := Detail[lLigne];
    lNum := lEcr.GetValue ('E_NUMLIGNE');
    InverserEcr (lEcr);
    if EstMultiEche (lNum) then
    begin
      lEChe := GetMultiEche (lNum);
      if lEche <> nil then
      for i:=0 to lEche.Detail.Count - 1 do
      begin
        lEcr := lEche.Detail[i];
        InverserEcr (lEcr);
      end;
    end;
  end;
end;

function ExtourneEng(
                      Eng: TPieceEngagement;          {Engagement à extourner}
                      EngExtourne: TPieceEngagement;  {Engagement extourné}
                      EngDate : TDateTime)            {Date de pièce extournée}
                      : boolean;
{ 1 étape
    Eng : prend le statut T avec la date Rappro à EngDate.
        Sauvegarde l'extension Eng
  2 étape
    Crée une pièce extourné avec
    Eng : DateComptable à EngDate
          Exercice selon EngDate
          Inverser les montants de toutes les lignes
          Statut engagement à X
          Sauvegarde la nouvelle pièce engagement avec extension.
  Eng est la pièce extournée
}
var
  vEngExt: TPieceEngagement;
begin
  Result := false;
  Eng.PutValueAll('CEN_STATUTENG','T');
  Eng.PutValueAll('CEN_DATERAPPRO', EngDate);
  if Eng.SaveEngage then
    begin
    {b FP 02/05/2006}
    if EngExtourne = nil then
      vEngExt := TPieceEngagement.CreerPieceEng
    else
      vEngExt := EngExtourne;
    try
      vEngExt.DupliquerPieceEng(Eng, False);
      vEngExt.PutEntete('E_DATECOMPTABLE', EngDate);
      vEngExt.PutEntete('E_EXERCICE', QuelExoDt(EngDate));
      vEngExt.PutEntete('E_NUMEROPIECE',0);
      vEngExt.InverserMontant;
      vEngExt.PutValueAll('CEN_STATUTENG','X');
      vEngExt.Action := taCreat;
      Result := vEngExt.Save;
    finally
      if vEngExt <> EngExtourne then
        vEngExt.Free;
      end;
    {e FP 02/05/2006}
    end;
end;

function ResteEng(Eng:      TPieceEngagement; {Pièce d'origine}
                  EngRest:  TPieceEngagement  {Piece contenant le reste}
                  ) : boolean;            {Eng reste à engager et écrite}
{ Eng est déjà rapproché avec une facture et écrit dans la base.
      A partir de l'écriture et son extension je calcule le reste à engager.
      S'il reste un pièce équilibré à générer alors
      je crée une pièce reste à engager avec le statut P
      en Sauvegardant la nouvelle pièce avec son extension.
  Eng est la pièce reste à engager
}

  Function ModifierMontant(Eng: TPieceEngagement; vLigne : integer):boolean;
  var
    MntR : double;
  begin
    if Eng.GetValue (vLigne, 'E_DEBITDEV') <> 0 then
    begin
      MntR := Eng.GetValue (vLigne, 'E_DEBITDEV') - Eng.GetValue (vLigne, 'CEN_MONTANTRAP');
      Eng.PutValue (vLigne, 'E_DEBITDEV', MntR);
    end
    else
    begin
      MntR := Eng.GetValue (vLigne, 'E_CREDITDEV') + Eng.GetValue (vLigne, 'CEN_MONTANTRAP');
      Eng.PutValue (vLigne, 'E_CREDITDEV', MntR);
    end;
    Result := true;
  end;

{FP 02/05/2006
var
  lOrig : TOB;
  lEEng : TOB;

  Procedure RecupererOrigine;
  var leTob,
      lxTob : TOB;
      i    : integer;
  begin
    Eng.ClearDetail;
    Eng.Dupliquer (lOrig,true,true);
    Eng.ViderTobExtension;
    Eng.FEngage.Dupliquer (lEEng,true,true);
    for i := 0 to Eng.Detail.Count -1 do
    begin
      leTob := Eng.Detail[i];
      CputTobEngage (leTob, nil);
    end;
    for i:=0 to Eng.FEngage.Detail.Count - 1 do
    begin
      lxTob := Eng.FEngage.Detail[i];
      if lxTob <> nil then
      begin
        leTob := Eng.GetTob (lxTob.GetValue ('CEN_NUMLIGNE'));
        if leTob <> nil then
          CPutTobEngage (leTob, lxTob);
      end;
    end;
  end;
}
  Function PieceOk(Eng: TPieceEngagement): boolean;
  var
    DCPiece  : RecCalcul;
  begin
    result := false;
    if Eng.Detail.Count > 1 then
    begin
//      if Eng.isValidPiece then begin
        DCPiece := CCalculSoldePiece(Eng);
        Result := (DCPiece.D  <> 0) and (DCPiece.D - DCPiece.C = 0);
//      end;
    end;
  end;

  Procedure PreparerPiece(Eng: TPieceEngagement);
  var
    lDate : TDateTime;
  begin
    lDate := Eng.GetValue (1,'CEN_DATERAPPRO');
    Eng.PutEntete ('E_DATECOMPTABLE', lDate);
    Eng.PutEntete ('E_EXERCICE', QuelExoDt(lDate));
    Eng.PutEntete ('E_NUMEROPIECE',0);
    Eng.PutValueAll('CEN_STATUTENG','P');
    Eng.PutValueAll('CEN_MONTANTRAP',0);
    Eng.PutValueAll('CEN_DATERAPPRO',iDate1900);
    Eng.PutValueAll('CEN_FDATECOMPTA',iDate1900);
    Eng.PutValueAll('CEN_FJOURNAL','');
    Eng.PutValueAll('CEN_FEXERCICE','');
    Eng.PutValueAll('CEN_FNUMEROPIECE',0);
    Eng.PutValueAll('CEN_FQUALIFPIECE','');
    Eng.PutValueAll('CEN_FNUMLIGNE',0);
    Eng.PutValueAll('CEN_FNUMECHE',0);
    Eng.RenumeroteLignes;
    Eng.Action := taCreat;
  end;

var
  i:         integer;
  NbrModif : integer;
  lNumLigne: integer;
  lSup:      integer;
  vEngRest:  TPieceEngagement;
begin   {ResteEng}
  Result := false;
  {b FP 02/05/2006 lOrig := Tob.Create ('sauve ligne et analytique',nil,-1);
  lOrig.Dupliquer (Eng, true, true);
  lEEng := Tob.Create ('sauve extension engage',nil,-1);
  lEEng.Dupliquer (Eng.FEngage,true,true);}

  if EngRest = nil then
    vEngRest := TPieceEngagement.CreerPieceEng
  else
    vEngRest := EngRest;
  try
    vEngRest.DupliquerPieceEng(Eng, False);

    lSup := 0;
    NbrModif := 0;
    for i:=vEngRest.Detail.Count downto 1 do
      begin
      lNumLigne := vEngRest.Detail[i-1].GetValue ('E_NUMLIGNE');
      if vEngRest.GetSoldeLigne(lNumLigne) = 0 then
        begin
        vEngRest.DeleteRecord(lNumLigne);
        inc(lSup);
        end
      else if ModifierMontant(vEngRest, lNumLigne) then
        inc(NbrModif);
      end;

    {b FP 02/05/2006
    if (lSup > 0) or (nbrModif > 0) then begin
      PreparerPiece;
      if PieceOk then begin
        if Eng.Save then
          Result := true
        else RecupererOrigine;
      end else RecupererOrigine;
    end;
    lOrig.Free;
    lEEng.Free;}
    if (lSup > 0) or (nbrModif > 0) then
      begin
      PreparerPiece(vEngRest);
      if PieceOk(vEngRest) then
        Result := vEngRest.Save;
      end;
  finally
    if vEngRest <> EngRest then
      vEngRest.Free;
    end;
  {e FP 02/05/2006}
end;

(*
Function RapprocherEngFac(
                       Eng,          {Piece engagement rapproché}
                       Fact: TPieceEngagement;    {Piece facture rapproché}
                       MajEng,       {extension engagement à écrire dans BD}
                       MajFac: boolean;{extension facture à écrire dans DB}
                       vDate : tDateTime) {date de rapprochement}
                       : boolean; {ok:rapprochement possible et correcte}
var
//  lEEcr : TOB;
  EngOrig,          {garde l'origine pour restorer si erreur}
  FacOrig : TOB;
  MntE  : double;
  Solde : double;     {solde du traitement doit être 0}
  NbrMnt: integer;    {nombre de ligne traité doit être >0}
  Res   : boolean;

  Procedure MajLigneEng (vEngLig:integer; vFacLig:integer; vMnt:double);
  var lMnt : double;
  begin
    inc (NbrMnt);
    Solde := Solde + vMnt;
    lMnt := Eng.GetValue(vEngLig, 'CEN_MONTANTRAP');
    Eng.PutValue(vEngLig, 'CEN_MONTANTRAP', vMnt + lMnt);
    Eng.PutValue(vEngLig, 'CEN_DATERAPPRO', vDate);
    Eng.PutValue(vEngLig, 'CEN_STATUTENG', 'T');
    Eng.PutValue(vEngLig, 'CEN_FJOURNAL', Fact.GetValue (vFacLig,'E_JOURNAL'));
    Eng.PutValue(vEngLig, 'CEN_FEXERCICE', Fact.GetValue (vFacLig,'E_EXERCICE'));
    Eng.PutValue(vEngLig, 'CEN_FNUMEROPIECE', Fact.GetValue (vFacLig,'E_NUMEROPIECE'));
    Eng.PutValue(vEngLig, 'CEN_FQUALIFPIECE', Fact.GetValue (vFacLig,'E_QUALIFPIECE'));
    Eng.PutValue(vEngLig, 'CEN_FNUMLIGNE', Fact.GetValue (vFacLig,'E_NUMLIGNE'));
    Eng.PutValue(vEngLig, 'CEN_FDATECOMPTA', Fact.GetValue (vFacLig,'E_DATECOMPTABLE'));
    Eng.PutValue(vEngLig, 'CEN_FNUMECHE', Fact.GetValue (vFacLig,'E_NUMECHE'));
    lMnt := Fact.GetValue(vFacLig, 'CEN_MONTANTRAP');
    Fact.PutValue(vFacLig, 'CEN_MONTANTRAP', vMnt + lMnt);
    Fact.PutValue(vFacLig, 'CEN_DATERAPPRO', vDate);
  end;

  Procedure RapprocherEngF (vEngLig:integer; ParCompte :boolean = true);
  {Rapprocher à partir d'une ligne Engagé à une facture}
  var
    MntEng : double;
    MntFac : double;
    lFacLig : integer;
    lRapproche    : boolean;
  begin
    for lFacLig:=1 to Fact.Detail.Count do
      begin
      MntEng := Eng.GetSoldeLigne(vEngLig);
      if MntEng <> 0 then
        begin
        if ParCompte then
          begin
          if Eng.GetValue (vEngLig,'E_TYPEMVT') = 'TTC' then
            lRapproche := Eng.GetValue (vEngLig,'E_AUXILIAIRE') = Fact.GetValue (lFacLig,'E_AUXILIAIRE')
          else
            lRapproche := Eng.GetValue (vEngLig,'E_GENERAL') = Fact.GetValue (lFacLig,'E_GENERAL');
          end
        else
          lRapproche := Eng.GetValue (vEngLig,'E_TYPEMVT') = Fact.GetValue (lFacLig,'E_TYPEMVT');
        if lRapproche then
          begin
          MntFac := Fact.GetSoldeLigne(lFacLig);
          if (MntFac <> 0) then
            begin
            if Abs (MntFac) >= Abs (MntEng) then
              MajLigneEng(vEngLig, lFacLig, MntEng)
            else
              MajLigneEng(vEngLig, lFacLig, MntFac);
            end;
          end;
        end;
      end;
  end;

  Function EcrireExtension : boolean;
  var Res : boolean;
  begin
    try
      begintrans;
      Res := true;
      if MajEng then
        Res := Res and Eng.SaveEngage;
      if MajFac then
        Res := Res and Fact.SaveEngage;
      if Res then
        Committrans
      else rollback;
    except
      on e : exception do
      begin
        Res := false;
        rollback;
        showmessage(e.message);
      end;
    end;
    Result := Res;
  end;

  Procedure ReprendreExtension;
  begin
    {b FP 18/04/2006}
    Eng.ViderTobExtension;
    Eng.FEngage.Dupliquer(EngOrig, true, true);

    Fact.ViderTobExtension;
    Fact.FEngage.Dupliquer(FacOrig, true, true);
    
    {e FP 18/04/2006}
  end;

var
  lELigne  : integer;
begin {RapprocherEngFac}
  Res   := false;
  Solde := 0;
  NbrMnt:= 0;
  EngOrig := Tob.Create ('ext engagement',nil,-1);
  FacOrig := Tob.Create ('ext facture',nil,-1);
  try
    EngOrig.Dupliquer(Eng.FEngage, true, true);
    FacOrig.Dupliquer(Fact.FEngage, true, true);
    for lELigne := 1 to Eng.Detail.Count do
    begin
      MntE  := Eng.GetSoldeLigne(lELigne);
      if MntE <> 0 then
        RapprocherEngF (lELigne, true); {sur compte}
    end;
    if (NbrMnt > 0) and (Solde <> 0) then {il y a des lignes non trouvées}
    begin
      for lELigne := 1 to Eng.Detail.Count do
      begin
        MntE  := Eng.GetSoldeLigne(lELigne);
        if MntE <> 0 then
          RapprocherEngF (lELigne, false); {sur type de ligne}
      end;
    end;
    if (NbrMnt > 0) and (Solde = 0) then begin {rapprochement réussit}
      Res := true;
      if MajEng or MajFac then
        Res := EcrireExtension;
      if not Res then
        ReprendreExtension;
    end
    else ReprendreExtension;
  finally
    Result := Res;
    EngOrig.Free;
    FacOrig.Free;
  end;
end;
*)

{b FP 02/05/2006}
Function RapprocherEngFac(
                       PieceEng,                       {Piece engagement rapproché}
                       PieceEngSimul,                  {Piece engagement simulé}
                       PieceFact: TPieceEngagement;    {Piece facture rapproché}
                       vDate : TDateTime)         {date de rapprochement}
                       : boolean;        {ok:rapprochement possible et correcte}
var
  Solde :  double;     {solde du traitement doit être 0}
  NbrMnt:  integer;    {nombre de ligne traité doit être >0}

  function OkRappro(TOB_1: TPieceCompta; Lig_1: Integer; TOB_2: TPieceCompta; Lig_2: Integer; ParCompte: boolean): Boolean;
  begin
    if ParCompte then
      begin
      if TOB_1.GetString(Lig_1,'E_TYPEMVT') = 'TTC' then
        Result := TOB_1.GetString(Lig_1,'E_AUXILIAIRE') = TOB_2.GetString(Lig_2,'E_AUXILIAIRE')
      else
        Result := TOB_1.GetString(Lig_1,'E_GENERAL') = TOB_2.GetString(Lig_2,'E_GENERAL');
      end
    else
      Result := TOB_1.GetString(Lig_1,'E_TYPEMVT') = TOB_2.GetString(Lig_2,'E_TYPEMVT');
  end;

  function RechercheFactureLigne(Eng: TPieceCompta; EngLig: Integer; ParCompte: boolean): Integer;
  var
    Lig: Integer;
  begin
    Result := 0;
    for Lig:=1 to PieceFact.Count do
      begin
        if OkRappro(Eng, EngLig, PieceFact, Lig, ParCompte) and
           (PieceFact.GetSoldeLigne(Lig) <> 0) then
          begin
          Result := Lig;
          break;
          end;
      end;
  end;

  function MajLigneEngFac(EngLig, EngSimulLig: integer; ParCompte :boolean; vMnt: double): Boolean;
  var
    lMnt :  double;
    NbDeci: Integer;
    FacLig: Integer;
  begin
    Result := False;
    NbDeci := PieceEng.Devise.Decimale;

     if PieceEngSimul = PieceFact then
      FacLig := EngSimulLig
    else
      begin
      FacLig := RechercheFactureLigne(PieceEngSimul, EngSimulLig, ParCompte);
      lMnt := PieceEngSimul.GetValue(EngSimulLig, 'CEN_MONTANTRAP');
      PieceEngSimul.PutValue(EngSimulLig, 'CEN_MONTANTRAP', Arrondi(vMnt+lMnt, NbDeci));
      end;

    if FacLig <> 0 then
      begin
      Result := True;

      lMnt := PieceEng.GetValue(EngLig, 'CEN_MONTANTRAP');
      PieceEng.PutValue(EngLig, 'CEN_MONTANTRAP',   Arrondi(vMnt+lMnt, NbDeci));
      PieceEng.PutValue(EngLig, 'CEN_DATERAPPRO',   vDate);
      PieceEng.PutValue(EngLig, 'CEN_STATUTENG',    'T');
      PieceEng.PutValue(EngLig, 'CEN_FJOURNAL',     PieceFact.GetString(FacLig,'E_JOURNAL'));
      PieceEng.PutValue(EngLig, 'CEN_FEXERCICE',    PieceFact.GetString(FacLig,'E_EXERCICE'));
      PieceEng.PutValue(EngLig, 'CEN_FNUMEROPIECE', PieceFact.GetInteger(FacLig,'E_NUMEROPIECE'));
      PieceEng.PutValue(EngLig, 'CEN_FQUALIFPIECE', PieceFact.GetString(FacLig,'E_QUALIFPIECE'));
      PieceEng.PutValue(EngLig, 'CEN_FNUMLIGNE',    PieceFact.GetInteger(FacLig,'E_NUMLIGNE'));
      PieceEng.PutValue(EngLig, 'CEN_FDATECOMPTA',  PieceFact.GetDateTime(FacLig,'E_DATECOMPTABLE'));
      PieceEng.PutValue(EngLig, 'CEN_FNUMECHE',     PieceFact.GetInteger(FacLig,'E_NUMECHE'));

      lMnt := PieceFact.GetValue(FacLig, 'CEN_MONTANTRAP');
      PieceFact.PutValue(FacLig, 'CEN_MONTANTRAP', Arrondi(vMnt+lMnt, NbDeci));
      PieceFact.PutValue(FacLig, 'CEN_DATERAPPRO', vDate);
      end;
  end;

  procedure RapprocherEng(EngLig:integer; ParCompte :boolean);
  {Rapprocher à partir d'une ligne Engagé à une facture}
  var
    MntEng:         double;
    MntEngSimul:    double;
    vMnt:           double;
    EngSimulLig:    integer;
    lRapproche:     boolean;
  begin
    for EngSimulLig:=1 to PieceEngSimul.Detail.Count do
      begin
      MntEng      := PieceEng.GetSoldeLigne(EngLig);
      MntEngSimul := PieceEngSimul.GetSoldeLigne(EngSimulLig);
      lRapproche  := False;
      if (MntEng <> 0) and (MntEngSimul <> 0) then
         lRapproche  := OkRappro(PieceEng, EngLig, PieceEngSimul, EngSimulLig, ParCompte);
      if lRapproche then
        begin
        if Abs(MntEngSimul) >= Abs(MntEng) then
          vMnt := MntEng
        else
          vMnt := MntEngSimul;
        if MajLigneEngFac(EngLig, EngSimulLig, ParCompte, vMnt) then
          begin
          inc(NbrMnt);
          Solde := Arrondi(Solde+vMnt, PieceEng.Devise.Decimale);
          end
        else
          break;
        end;
      end;
  end;

  Function EcrireExtension: boolean;
  begin
    Result := True;
    begintrans;
    try
      Result := Result and PieceEng.SaveEngage;
      Result := Result and PieceFact.SaveEngage;
      if Result then
        Committrans
      else
        rollback;
    except on E: Exception do
      begin
        Result := False;
        rollback;
        showmessage(e.message);
      end;
    end;
  end;

var
  MntE  :       double;
  lELigne:      integer;
  EngOrig:      Tob;
  EngSimulOrig: Tob;
  FacOrig:      Tob;
begin          {RapprocherEngFac}
  Result := False;
  Solde  := 0;
  NbrMnt := 0;

  EngOrig      := Tob.Create ('ext eng',nil,-1);
  EngSimulOrig := Tob.Create ('ext eng simu',nil,-1);
  FacOrig      := Tob.Create ('ext fact',nil,-1);
  try
    EngOrig.Dupliquer(PieceEng.FEngage, true, true);
    EngSimulOrig.Dupliquer(PieceEngSimul.FEngage, true, true);
    FacOrig.Dupliquer(PieceFact.FEngage, true, true);

    for lELigne := 1 to PieceEng.Detail.Count do
      begin
      MntE  := PieceEng.GetSoldeLigne(lELigne);
      if MntE <> 0 then
        RapprocherEng(lELigne, true); {sur compte}
      end;
    if (NbrMnt > 0) and (Solde <> 0) then {il y a des lignes non trouvées}
      begin
      for lELigne := 1 to PieceEng.Detail.Count do
        begin
        MntE  := PieceEng.GetSoldeLigne(lELigne);
        if MntE <> 0 then
          RapprocherEng(lELigne, false); {sur type de ligne}
        end;
      end;
    if (NbrMnt > 0) and (Solde = 0) then {rapprochement réussit}
      Result := EcrireExtension;
    if not Result then
      begin
      PieceEng.ViderTobExtension;
      PieceEng.FEngage.Dupliquer(EngOrig, true, true);
      PieceEngSimul.ViderTobExtension;
      PieceEngSimul.FEngage.Dupliquer(EngSimulOrig, true, true);
      PieceFact.ViderTobExtension;
      PieceFact.FEngage.Dupliquer(FacOrig, true, true);
      end;
  finally
    EngOrig.Free;
    EngSimulOrig.Free;
    FacOrig.Free;
    end;
end;
{e FP 02/05/2006}

function SupprimerEngagement (Eng : tPieceEngagement):boolean;
         {supprimer un engagement de la base}
var
    lStat: string;
    lTobO: Tob;
    Meme : boolean;
    lReq : string;
begin
  Result:=false;
  lStat := Eng.GetValue (1,'CEN_STATUTENG');
  if length(lStat) < 1 then lStat := ' ';
  if (Eng.GetEntete ('E_QUALIFPIECE')='P') and
     (lStat[1] in ['E','P','T','X']) then
  begin
    lTobO := Eng.TobOrigine;
    if lTobO.Detail.Count > 1 then
      Meme := (Eng.GetEntete ('E_JOURNAL')   = lTobO.Detail[0].GetValue ('E_JOURNAL'))
         and  (Eng.GetEntete ('E_EXERCICE')  = lTobO.Detail[0].GetValue ('E_EXERCICE'))
         and  (Eng.GetEntete ('E_DATECOMPTABLE')=lTobO.Detail[0].GetValue ('E_DATECOMPTABLE'))
         and  (Eng.GetEntete ('E_NUMEROPIECE')=lTobO.Detail[0].GetValue ('E_NUMEROPIECE'))
         and  (Eng.GetEntete ('E_QUALIFPIECE')=lTobO.Detail[0].GetValue ('E_QUALIFPIECE'))
    else Meme := false;
    if not Meme then begin
      lReq := 'E_JOURNAL="'            + Eng.GetEntete('E_JOURNAL') + '"'
            + ' AND E_EXERCICE="'      + Eng.GetEntete('E_EXERCICE') + '"'
            + ' AND E_DATECOMPTABLE="' + UsDateTime( Eng.GetEntete('E_DATECOMPTABLE') ) + '"'
            + ' AND E_NUMEROPIECE='    + InttoStr( Eng.GetEntete('E_NUMEROPIECE') )
            + ' AND E_QUALIFPIECE="'   + Eng.GetEntete('E_QUALIFPIECE') + '"';
      Eng.LoadFromSQL (lReq);
    end;
    Eng.ViderDBExtension;
    Eng.DetruitPiece;
    Result := true;
  end;
end;

function DerapprocherEngagement ( Eng : tPieceEngagement ) :boolean;
         {Ramener à l'engagement initial en
            supprimant les engagements extournés et les restes
            maj les montants rapprochés avec les factures}


  Function SupprimerLienFacture (vEng : tPieceEngagement):boolean;
  {Engagement avec Statut T doit être rapproché avec une facture
    1 - Si la facture est rapproché avec 1 seul engagement
          alors je supprime simplement le lien
        Sinon je mets à jour le montant rapproché
    2 - Si c'était un engagement initial
          alors je remet en E
        Sinon je supprime}

    Function ReqFacture (vLig:integer) : string;
    begin
      Result := 'E_JOURNAL = "'  + vEng.GetValue (vLig, 'CEN_FJOURNAL') + '"' +
             ' AND E_EXERCICE="' + vEng.GetValue (vLig, 'CEN_FEXERCICE') + '"'+
             ' AND E_DATECOMPTABLE="'+UsDateTime (vEng.GetValue (vLig, 'CEN_FDATECOMPTA'))+'"'+
             ' AND E_NUMEROPIECE='+ IntToStr (vEng.GetValue (vLig, 'CEN_FNUMEROPIECE'))+
             ' AND E_QUALIFPIECE="' + vEng.GetValue (vLig, 'CEN_FQUALIFPIECE') + '"';
    end;

    Function SupprimerExtensionFacture (vLig:integer) : boolean;
    var lFac : tPieceEngagement;
    begin
      lFac := TPieceEngagement.CreerPieceEng;
      lFac.LoadFromSQL( ReqFacture(vLig) );
      lFac.ViderDBExtension;
      lFac.Free;
      Result := true;
    end;

    Function MajExtensionFacture (vLig:integer) : boolean;
    var lFac : tPieceEngagement;
        lLig : integer;
        lSolde   : double;
        lNbrMnt  : integer;

      Procedure DefaireFacture (Lig:integer; ParCompte:boolean);
      var i   : integer;
         MntF : double;
         MntE : double;
         lRap : boolean;
      begin
        for i:=1 to lFac.Detail.Count do
        begin
          MntE := Arrondi (vEng.GetValue (Lig, 'CEN_MONTANTRAP'), vEng.Devise.Decimale);
          if MntE <> 0 then
          begin
            MntF := Arrondi (lFac.GetValue (i, 'CEN_MONTANTRAP'), lFac.Devise.Decimale);
            if MntF <> 0 then
            begin
              if ParCompte then
                begin
                if vEng.GetValue (Lig, 'E_TYPEMVT') = 'TTC' then
                  lRap := vEng.GetValue (Lig, 'E_AUXILIAIRE') = lFac.GetValue (i, 'E_AUXILIAIRE')
                else
                  lRap := vEng.GetValue (Lig, 'E_GENERAL') = lFac.GetValue (i, 'E_GENERAL');
                end
              else
                lRap := vEng.GetValue (Lig, 'E_TYPEMVT') = lFac.GetValue (i, 'E_TYPEMVT');
              if lRap then
              begin
                if Abs(MntF) > Abs (MntE) then
                begin
                  vEng.PutValue (Lig, 'CEN_MONTANTRAP', 0);
                  lFac.PutValue (i, 'CEN_MONTANTRAP', MntF-MntE);
                  lSolde := lSolde + MntE;
                end
                else
                begin
                  vEng.PutValue (Lig, 'CEN_MONTANTRAP', MntE-MntF);
                  lFac.PutValue (i, 'CEN_MONTANTRAP', 0);
                  lSolde := lSolde + MntF;
                end;
                inc (lNbrMnt);
              end;
            end;
          end;
        end;
      end;

    begin {MajExtensionFacture}
      Result := false;
      lFac := TPieceEngagement.CreerPieceEng;
      lFac.LoadFromSQL (ReqFacture (vLig));
      lSolde := 0;
      lNbrMnt:= 0;
      for lLig := 1 to vEng.Detail.Count do
      if Arrondi (vEng.GetValue (lLig, 'CEN_MONTANTRAP'),
                  vEng.Devise.Decimale) <> 0 then
        DefaireFacture (lLig, true);
      for lLig := 1 to vEng.Detail.Count do
      if Arrondi (vEng.GetValue (lLig, 'CEN_MONTANTRAP'),
                  vEng.Devise.Decimale) <> 0 then
        DefaireFacture (lLig, false);
      if (lNbrMnt > 0) and (lSolde = 0) then
        Result := lFac.SaveEngage;
      lFac.Free;
    end;

  var lQ : tQuery;
      lLig : integer;
      lReq : string;
  begin {SupprimerLienFacture}
    Result := true;
    lLig := 1;
    while (vEng.GetValue (lLig, 'CEN_FJOURNAL') = '') and not (vEng.IsOut(lLig)) do
      inc (lLig);
    if (vEng.GetValue (lLig, 'CEN_FJOURNAL') <> '') and (not vEng.IsOut(lLig)) then
    begin
      lReq := 'SELECT * FROM CENGAGEMENT '+
          ' WHERE CEN_FJOURNAL="'+ vEng.GetValue (lLig, 'CEN_FJOURNAL')+'"'+
          '   AND CEN_FEXERCICE="' + vEng.GetValue (lLig, 'CEN_FEXERCICE')+'"'+
          '   AND CEN_FDATECOMPTA="' + UsDateTime (vEng.GetValue (lLig, 'CEN_FDATECOMPTA'))+'"'+
          '   AND CEN_FNUMEROPIECE=' + IntToStr (vEng.GetValue (lLig, 'CEN_FNUMEROPIECE'))+
          '   AND CEN_FQUALIFPIECE="' + vEng.GetValue (lLig, 'CEN_FQUALIFPIECE')+'"'+
          '   AND CEN_EJOURNAL<>"'+ vEng.GetValue (lLig, 'CEN_EJOURNAL')+'"'+
          '   AND CEN_EEXERCICE<>"' + vEng.GetValue (lLig, 'CEN_EEXERCICE')+'"'+
          '   AND CEN_EDATECOMPTA<>"' + UsDateTime (vEng.GetValue (lLig, 'CEN_EDATECOMPTA'))+'"'+
          '   AND CEN_ENUMEROPIECE<>' + IntToStr (vEng.GetValue (lLig, 'CEN_ENUMEROPIECE'))+
          '   AND CEN_EQUALIFPIECE<>"' + vEng.GetValue (lLig, 'CEN_EQUALIFPIECE')+'"';
      lQ := OpenSelect(lReq, vEng.Dossier);
      if lQ.Eof then
        Result := SupprimerExtensionFacture (lLig)
      else
        Result := MajExtensionFacture (lLig);
      Ferme(lQ);
    end;

    {Dans le cas d'un rappro terminé, il n'y a pas de facture associée à l'engagement}
    if Result then
    begin
      if vEng.IsEngagementInitial then
      begin
        vEng.ViderTobExtension;
        Result := vEng.SaveEngage;
      end
      else
        Result := SupprimerEngagement(vEng);
    end;

  end;

var
  lReqEng    : string; {engagement initial}
  lQEngInit  : tQuery;
  lEng       : tPieceEngagement;
  lNbrEng    : integer;
  lOk        : boolean;
begin {DerapprocherEngagement}
  Result := false;
  lNbrEng:= 0;
  lQEngInit := OpenSelect('SELECT CEN_JOURNAL, CEN_EXERCICE, '
        + ' CEN_DATECOMPTABLE, CEN_NUMEROPIECE, CEN_QUALIFPIECE FROM CENGAGEMENT '
        + ' WHERE CEN_EJOURNAL="'      + Eng.GetValue(1,'CEN_EJOURNAL') + '"'
        + ' AND CEN_EEXERCICE="'       + Eng.GetValue(1,'CEN_EEXERCICE') + '"'
        + ' AND CEN_EDATECOMPTA="'  + UsDateTime( Eng.GetValue(1,'CEN_EDATECOMPTA') ) + '"'
        + ' AND CEN_ENUMEROPIECE='     + InttoStr( Eng.GetValue(1,'CEN_ENUMEROPIECE') )
        + ' AND CEN_EQUALIFPIECE="'    + Eng.GetValue(1,'CEN_EQUALIFPIECE') + '"'
        + ' GROUP BY CEN_JOURNAL, CEN_EXERCICE, CEN_DATECOMPTABLE, '
        + ' CEN_NUMEROPIECE, CEN_QUALIFPIECE '
        + ' ORDER BY CEN_JOURNAL, CEN_EXERCICE, CEN_DATECOMPTABLE, '
        + ' CEN_NUMEROPIECE, CEN_QUALIFPIECE ', Eng.Dossier ) ;
  lEng := tPieceEngagement.CreerPieceEng;
  try
    begintrans;
    lOK := true;
    While (not lQEngInit.Eof) and lOk do
    begin
      lReqEng := 'E_JOURNAL = "'  + lQEngInit.FindField ('CEN_JOURNAL').asString+ '"' +
         ' AND E_EXERCICE="' + lQEngInit.FindField ('CEN_EXERCICE').asString+ '"'+
         ' AND E_DATECOMPTABLE="'+UsDateTime (lQEngInit.FindField ('CEN_DATECOMPTABLE').asDateTime)+'"'+
         ' AND E_NUMEROPIECE='+ IntToStr (lQEngInit.FindField ('CEN_NUMEROPIECE').asInteger)+
         ' AND E_QUALIFPIECE="' + lQEngInit.FindField ('CEN_QUALIFPIECE').asString + '"';
      lEng.LoadFromSQL (lReqEng);
      if lEng.GetValue (1,'CEN_STATUTENG') = 'T' then
        lOk := SupprimerLienFacture(lEng)
      else
        lOk := SupprimerEngagement(lEng);
      if lOk then
        inc(lNbrEng);
      lQEngInit.Next;
    end;
    if (lNbrEng > 0) and lOk then begin
      committrans;
      Result:=true;
    end
    else
      rollback;
  except
    on e : exception do
    begin
      rollback;
      showmessage(e.message);
    end;
  end;
  lEng.Free;
  Ferme(lQEngInit);
end;

Function Rapprocher (sEng, sFac : string):integer;
var
  lEng : TPieceEngagement;
  lFac : TPieceEngagement;
  lDate: TDateTime;
begin
  Result := 0;
  lDate := Date;
  lEng := TPieceEngagement.CreerPieceEng;
  lFac := TPieceEngagement.CreerPieceEng;
  BeginTrans;
  try
    lEng.LoadFromSQL (sEng);
    if ExtourneEng (lEng, nil, lDate) then
    begin
      Result := 1;
      lEng.LoadFromSQL (sEng);
      lFac.LoadFromSQL (sFac);
      if RapprocherEngFac(lEng, lFac, lFac, lDate) then
      begin
        inc (Result);
        if lEng.IsEngagementEstTermine then
          CommitTrans
        else if ResteEng (lEng, nil) then
        begin
          CommitTrans;
          inc (Result);
        end
        else begin
          Rollback;
          Result := -Result;
        end;
      end
      else begin
        Rollback;
        Result := -Result;
      end;
    end
    else begin
      rollback;
      Result := -Result;
    end;
  except
    rollback;
    Result := -Result;
  end;
  lEng.Free;
  lFac.Free;
end;

Procedure EssaiEngagement;

  Procedure RapprocherEF;
  var ReqE,
      ReqF : string;
      Res : integer;
  begin
    ReqE := 'E_JOURNAL="ACH" AND E_NUMEROPIECE=247 AND E_EXERCICE="007"';
    ReqF := 'E_JOURNAL="ACH" AND E_NUMEROPIECE=674 AND E_EXERCICE="007"';
    Res  := 0;
    if InputQuery ('Rapprochement Engagement Facture','Piece Engagement: ',ReqE) then
    if InputQuery ('Rapprochement Engagement Facture','Piece Facture: ',ReqF) then
    Res := Rapprocher (ReqE, ReqF);
    Case Res of
      1 : ReqE := 'Extourné';
      2 : ReqE := 'Extourné + Rapproché';
      3 : ReqE := 'Extourné + Rapproché + Reste';
      -1: ReqE := 'Erreur Rapproché, mais Extourné Ok';
      -2: ReqE := 'Erreur Reste, mais ok sur Ext et Rap';
      -3: ReqE := 'Erreur Impossible';
      else ReqE:= 'Inconnu';
    end;
    ShowMessage (ReqE+' valeur = '+IntToStr(Res));
  end;

  procedure DeRapprocherEF;
  var Req : string;
      Eng : tPieceEngagement;
  begin
    Req := 'E_JOURNAL="ACH" AND E_NUMEROPIECE=247 AND E_EXERCICE="007"';
    if InputQuery ('DeRapprochement Engagement Facture','Piece Engagement: ',Req) then
    begin
      Eng := tPieceEngagement.CreerPieceEng;
      Eng.LoadFromSQL (Req);
      if DerapprocherEngagement (Eng) then
        ShowMessage ('Dérapprocher Ok')
      else ShowMessage ('Erreur');
      Eng.Free;
    end;
  end;

var
  Rep : string;
begin
  Rep := 'R';
  if InputQuery ('Test outil Rappro ou Dérappro Engagement--Facture',
              'Rappro=R, Derappro=D ',Rep) then begin
    if Length(Rep)<1 then Rep:=' ';
    Case Upcase(Rep[1]) of
      'R' : RapprocherEF;
      'D' : DerapprocherEF;
    end;
  end;
end;

{b FP 02/05/2006}
procedure TPieceEngagement.DupliquerPieceEng(PieceSrc: TPieceEngagement; RAZExtension: Boolean);
var
  i: Integer;
begin
  inherited DupliquerPiece(PieceSrc);

  ViderTobExtension;
  if not RAZExtension then
    begin
    FEngage.Dupliquer(PieceSrc.FEngage, True, True);
    for i:=0 to FEngage.Detail.Count-1 do
      CPutTobEngage(Detail[i], FEngage.Detail[i]);
    MajCleInit;
    end;
end;

procedure TPieceEngagement.DupliquerPiece(PieceSrc: TPieceCompta);
begin
  if PieceSrc is TPieceEngagement then
    DupliquerPieceEng( (PieceSrc as TPieceEngagement), True)
  else
    inherited;
end;
{e FP 02/05/2006}

procedure TPieceEngagement.CFreeExtension(vTobEcr: Tob);
begin
  inherited CFreeExtension(vTobEcr);
  CFreeTOBEngage( vTobEcr ) ; // libération de l'Engagement
end;

procedure TPieceEngagement.SynchroniseTobEngage;
var lTobTmp   : Tob ;
    i         : integer ;
    lTobEng : Tob ;
begin

  if FEngage.Detail.count = 0 then Exit ;

  lTobTmp := Tob.Create('$TEMP', nil, -1);
  // On isole les TobCompl ayant une référence sur les lignes
  for i := 0 to Detail.count-1 do
    begin
    lTobEng := CGetTobEngage( Detail[i] ) ;
    if Assigned( lTobEng ) then
      lTobEng.ChangeParent( lTobTmp, -1 ) ;
    end ;

  // efface les Tob "Engagement" isolées
  FEngage.ClearDetail ;

  // réintégration des tob Engage
  for i := lTobTmp.Detail.count-1 downto 0 do
    begin
    lTobEng := lTobTmp.Detail[i] ;
    lTobEng.ChangeParent( FEngage, -1 ) ;
    end ;

end;

procedure TPieceEngagement.DupliquerLigne(vNumLigneDst: Integer;  PieceSrc: TPieceCompta; vNumLigneSrc: Integer);
var lTob : TOB ;
begin
  inherited DupliquerLigne(vNumLigneDst,PieceSrc,vNumLigneSrc);
  lTob := CGetTob( vNumLigneDst ) ;
  if lTob <> nil then
    CPutTobEngage(lTob,nil);
//  if lTob.FieldExists('PENGAGE') then
//    lTob.PutValue('PENGAGE',-1) ;
end;

function TPieceEngagement.GetMontantRAP(vNumLigne: Integer): double;
var lTObEcr : Tob ;
    lTobEng : Tob ;
begin
  result := 0 ;

  lTobEcr := CGetTOB(vNumLigne);
  if lTobEcr<>nil then
    begin
    lTobEng := CGetTOBEngage(lTobEcr) ;
    if lTobEng <> nil then
      Result := lTobEng.GetDouble('CEN_MONTANTRAP') ;
    end ;
end;

end.
