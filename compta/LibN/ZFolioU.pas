unit ZFolioU;

interface

uses {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     HCtrls,
     SaisUtil, Ent1, UTOB,                                          { +- COMPTA  -+ }
     TZ, ZTypes, ZCompte, ZTiers, ZFolio;                      { +- PFUGIER -+ }

type
  TFolioU = class(TObject)
  private
  FObjFolio:  TFolio ;
  FNumFolio:  string ;
  FCodeJal:   string ;
  FbJalLibre: Boolean ;
  FExo:       string ;
  FDateFin:   TDateTime ;
  FEtabl:     string ;
  FmemPivot:  RDevise ;
  FmemOppos:  RDevise ;
//  FmemDevise: RDevise ;
  FComptes:   TZCompte ;
  FTiers:     TZTiers ;
  FmemJal:    RJal ;
  function    GetEcrs: TOB;
  procedure   InitMonnaie ;
  procedure   ChargeJal ;
//  function    GetAnySolde(Piece : LongInt; var SP, SD : Double) : LongInt ;
//  function    SoldeLaLigne(Piece, DefRow : LongInt; var SP, SD : Double) : Boolean ;
//  function    GetRowHT(Piece, DefRow : LongInt) : LongInt ;
//  function    GetRowTiers(Piece, DefRow : LongInt) : LongInt ;
//  function    IsRowHT(Row : LongInt) : Boolean ;
//  function    IsRowTiers(Row : LongInt) : Boolean ;
//  procedure   AjusteLigneEuro(Ecrs : TZF) ;
  procedure   SetNum(Row : LongInt) ;
//  procedure   SetPiece ;
//  function    GetRowLast(Piece : LongInt) : LongInt ;
//  function    GetRowFirst(Piece : LongInt) : LongInt ;
//  function    PieceToTZF(Piece : LongInt) : TZF ;
//  function    TZFToPiece(Piece : LongInt; Source : TZF) : Boolean ;
//  procedure   CreateRow(Row : LongInt; bRowEcart : Boolean=FALSE) ;
//  function    AfterCreateRow(Row: LongInt) : Boolean ;
//  procedure   SetRowLib(Row : LongInt) ;
//  function    AjouteLigneEcart(LastRow : LongInt; EcartP, EcartE : Double) : Boolean ;
//  function    VerifEquilibre(Piece, DefRow : LongInt) : Boolean ;
  function    GetOptions: ROpt ;
  protected
  public
  constructor Create(NumFolio, CodeJal, Exo: string; DateFin: TDateTime; bJalLibre: Boolean);
  destructor  Destroy; override;
  function    Read: Boolean;
  function    Write: Boolean;
//  function    VerifFolio : Boolean ;
//  function    VerifOppose : Boolean ;
  function    VerifNumLigne : Boolean ;
  property    NumFolio:  string    read FNumFolio  write FNumFolio;
  property    CodeJal:   string    read FCodeJal   write FCodeJal;
  property    Exo:       string    read FExo       write FExo;
  property    DateFin:   TDateTime read FDateFin   write FDateFin;
  property    bJalLibre: Boolean   read FbJalLibre write FbJalLibre;
  property    Ecrs:      TOB       read GetEcrs ;
  end ;

implementation

uses SysUtils,                                                 { +- Delphi  -+ }
     Hent1,                                                    { +- AGL     -+ }
     SaisComm ;                                      { +- COMPTA  -+ }
                                                               { +- PFUGIER -+ }

constructor TFolioU.Create(NumFolio, CodeJal, Exo: string; DateFin: TDateTime; bJalLibre: Boolean);
begin
FObjFolio:=nil;
FNumFolio:=NumFolio;
FCodeJal:=CodeJal;
FExo:=Exo;
FDateFin:=DateFin;
FbJalLibre:=bJalLibre;
FComptes:=TZCompte.Create() ;
FTiers:=TZTiers.Create ;
InitMonnaie;
end ;

destructor TFolioU.Destroy;
begin
inherited Destroy;
if FObjFolio<>nil then FObjFolio.Free;
if FComptes<>nil  then FComptes.Free ;
if FTiers<>nil    then FTiers.Free ;
end ;

function TFolioU.GetEcrs: TOB;
begin
Result:=FObjFolio.Ecrs;
end ;

function TFolioU.Read: Boolean;
var Jal: RJal; Opt: ROpt; TotSavDebit, TotSavCredit : Double;
begin
{ Strict minimum pour créer/charger un bordereau }
Jal.Code:=FCodeJal;
Opt:=GetOptions ;
FObjFolio:=TFolio.Create(FNumFolio,Jal, Opt) ;
Result:=FObjFolio.Read(FALSE, TotSavDebit, TotSavCredit, FALSE, FALSE, FbJalLibre);
FEtabl:=VH^.EtablisDefaut ;
if (Result) and (FObjFolio.Ecrs.Detail.Count>0)
  then FEtabl:=FObjFolio.Ecrs.Detail[0].GetValue('E_ETABLISSEMENT') ;
end;

function TFolioU.GetOptions: ROpt ;
var Opt: ROpt; Year, Month, Day: Word;
begin
Opt.Exo:=FExo;
DecodeDate(FDateFin, Year, Month, Day);
Opt.MaxJour:=DaysPerMonth(Year, Month);
Opt.Year:=Year;
Opt.Month:=Month;
if Opt.Exo=VH^.Encours.Code then Opt.TypeExo:=teEncours else
  if Opt.Exo=VH^.Suivant.Code then Opt.TypeExo:=teSuivant else
    Opt.TypeExo:=tePrecedent ;
Result:=Opt ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 18/07/2002
Modifié le ... :   /  /
Description .. : suppression d'un warning
Mots clefs ... :
*****************************************************************}
function TFolioU.Write: Boolean;
var Opt: ROpt;
begin
Result:=FALSE;
if FObjFolio=nil then Exit;
{ Strict minimum pour sauvegarde un bordereau }
ChargeJal ;
Opt:=GetOptions ;
{ Pas de mise à jour sur un exercice fermé ! }
if (Opt.TypeExo<>teEncours) and (Opt.TypeExo<>teSuivant) then Exit ;
//Result:=FObjFolio.Del(FmemJal) ;
FObjFolio.Del(FmemJal) ;
FmemJal.LastDateSais:=EncodeDate(Opt.Year, Opt.Month, Opt.MaxJour);
FmemJal.TotFolDebit:=FObjFolio.Ecrs.Somme( 'E_DEBIT',  ['E_NUMEROPIECE'], [FNumFolio], TRUE);
FmemJal.TotFolCredit:=FObjFolio.Ecrs.Somme('E_CREDIT', ['E_NUMEROPIECE'], [FNumFolio], TRUE);
Result:=FObjFolio.Write(FmemJal);
end;

procedure TFolioU.InitMonnaie ;
begin
FmemPivot.Code:=V_PGI.DevisePivot ;
if FmemPivot.Code<>'' then GetInfosDevise(FmemPivot) ;
FmemOppos.Code:=V_PGI.DeviseFongible ;
if FmemOppos.Code<>'' then GetInfosDevise(FmemOppos) ;
end ;

procedure TFolioU.ChargeJal ;
var Q : TQuery ; DateMvt, DebPeriode, FinPeriode : TDateTime ; St : string ; Opt : ROpt ;
begin
if FCodeJal='' then Exit ;
Q:=OpenSQL('SELECT J_JOURNAL, J_NATUREJAL, J_COMPTEINTERDIT, J_MULTIDEVISE, J_CONTREPARTIE,'
          +' J_TOTDEBP, J_TOTCREP, J_TOTDEBE, J_TOTCREE, J_TOTDEBS, J_TOTCRES, J_MODESAISIE,'
          +' J_VALIDEEN, J_VALIDEEN1, J_DATEDERNMVT, J_NUMDERNMVT, J_COMPTEAUTOMAT'
          +' FROM JOURNAL WHERE J_JOURNAL="'+FCodeJal+'"', TRUE) ;
FmemJal.Code:=FCodeJal ;
FmemJal.Nature:=Q.FindField('J_NATUREJAL').AsString ;
FmemJal.CInterdit:=Q.FindField('J_COMPTEINTERDIT').AsString ;
FmemJal.CAutomat:=Q.FindField('J_COMPTEAUTOMAT').AsString ;
FmemJal.bDevise:=(Q.FindField('J_MULTIDEVISE').AsString='X') ;
FmemJal.ValideN:=Q.FindField('J_VALIDEEN').AsString ;
FmemJal.ValideN1:=Q.FindField('J_VALIDEEN1').AsString ;
FmemJal.CContreP:=Q.FindField('J_CONTREPARTIE').AsString ;
FmemJal.ModeSaisie:=Q.FindField('J_MODESAISIE').AsString ;
if (FmemJal.Nature='BQE') or (FmemJal.Nature='CAI') then FmemJal.CAutomat:=FmemJal.CContreP+';' ;
FmemJal.NbAuto:=0 ; FmemJal.CurAuto:=0 ;
St:=FmemJal.CAutomat ; while ReadTokenSt(St)<>'' do Inc(FmemJal.NbAuto) ;
Opt:=GetOptions ;
DebPeriode:=EncodeDate(Opt.Year, Opt.Month, 1) ;
FinPeriode:=EncodeDate(Opt.Year, Opt.Month, Opt.MaxJour) ;
DateMvt:=Q.FindField('J_DATEDERNMVT').AsDateTime ;
if (DateMvt>=DebPeriode) and (DateMvt<=FinPeriode)
   then FmemJal.LastNum:=Q.FindField('J_NUMDERNMVT').AsInteger
   else FmemJal.LastNum:=1 ;
{if VH^.Entree.Code=VH^.Precedent.Code then
   begin
   FmemJal.TotExoDebit:=Q.FindField('J_TOTDEBP').AsFloat ;
   FmemJal.TotExoCredit:=Q.FindField('J_TOTCREP').AsFloat ;
   end ;
if VH^.Entree.Code=VH^.Encours.Code then
   begin
   FmemJal.TotExoDebit:=Q.FindField('J_TOTDEBE').AsFloat ;
   FmemJal.TotExoCredit:=Q.FindField('J_TOTCREE').AsFloat ;
   end ;
if VH^.Entree.Code=VH^.Suivant.Code then
   begin
   FmemJal.TotExoDebit:=Q.FindField('J_TOTDEBS').AsFloat ;
   FmemJal.TotExoCredit:=Q.FindField('J_TOTCRES').AsFloat ;
   end ;  }
FmemJal.TotFolDebit:=0 ; FmemJal.TotFolCredit:=0 ;
FmemJal.TotSaiDebit:=0 ; FmemJal.TotSaiCredit:=0 ;
FmemJal.TotVirDebit:=0 ; FmemJal.TotVirCredit:=0 ;
FmemJal.TotDebDebit:=0 ; FmemJal.TotDebCredit:=0 ;
Ferme(Q) ;
// Totaux sur la période
Q:=OpenSQL('SELECT SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE'
          +' WHERE E_JOURNAL="'+FCodeJal+'"'
          +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Opt.Year, Opt.Month, 1))+'"'
          +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(Opt.Year, Opt.Month, Opt.MaxJour))+'"'
          +' AND E_EQUILIBRE<>"X"',
          TRUE) ;
if (Q.BOF) and (Q.EOF) then
   begin
   FmemJal.TotPerDebit:=0 ; FmemJal.TotPerCredit:=0 ;
   end else
   begin
   FmemJal.TotPerDebit:=Q.Fields[0].AsFloat ;
   FmemJal.TotPerCredit:=Q.Fields[1].AsFloat ;
   end ;
Ferme(Q) ;
end ;

procedure TFolioU.SetNum(Row : LongInt) ;
var i : Integer ;
begin
for i:=0 to FObjFolio.Ecrs.Detail.Count-1 do FObjFolio.Ecrs.Detail[i].PutValue('E_NUMLIGNE', i+1) ;
end ;


function TFolioU.VerifNumLigne : Boolean ;
var i : integer ;
begin
Result:=FALSE ;
for i:=0 to FObjFolio.Ecrs.Detail.Count-1 do
  if FObjFolio.Ecrs.Detail[i].GetValue('E_NUMLIGNE')<>i+1
    then begin SetNum(0) ; Result:=TRUE ; Break ; end ;
end ;


end.
