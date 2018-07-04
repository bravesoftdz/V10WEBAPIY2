unit CalculSuggestion;

interface

uses SysUtils, Classes, UTOB, Hctrls, HStatus, FactUtil, StdCtrls, math,
     EntGC, Ent1, SaisUtil, FactComm, FactCalc, UTofListeInv,AglInit,UtilPGI, FactTOB, FactPiece, TiersUtil,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     BtpUtil,
     BTPrepaLivr,
     UtilSuggestion,
     UPlannifChUtil,
{$ENDIF}
     ParamSoc, HEnt1,UtilArticle,SelectDocSuggestion,uEntCommun;

type IrpMode = (IrpRef,IrpPrec,IrpOrig);

function Calcul_Suggestion(Params : string; var TobPiece : TOB; var TOBLDEVCHAR : TOB) : boolean;
procedure TraiteUnArticle(var TobPiece : TOB; TOBArt : Tob);
procedure CalculHistoStock(TOBArt : TOB; var TobMvts : TOB; LigneStd : boolean = true);
procedure CalculStockPhysique(TOBArt : TOB; var TobEvol : TOB);
procedure CalculEvolutionStock(TOBArt, TobMvts : TOB; var TobEvol : TOB;LigneStd:boolean=true;TOBLTraite : TOB = nil;FromCalculDocuments : boolean=false; Libart : String='');
procedure CalculRupture(var TobPiece ,TOBArt, TobMvts : TOB; var TobEvol : TOB);
procedure EnregDonnees(var TobPiece : TOB; TOBArt, TobEvol, TobMvts : TOB;TOBL : TOB;Lignestd:boolean=true;FromCalculDocuments : boolean=false);
// --
procedure EclateDate(DateOrig : TDateTime; var Annee, Mois, Jour : string);
function  CreerFilleEvol(TobEvol : TOB; Depot, Annee, Mois : string) : TOB;
procedure GenereEvolPeriode(TobEvol : TOB; Depot : string);
function  CreerFilleMvts(TobMvts : TOB; Depot, Annee, Mois : string) : TOB;
procedure GenereMvtsPeriode(TobMvts : TOB; Depot : string);
procedure Calcul_Fournisseur(TobSuggest : TOB; Param : string);
procedure Recherche_Tarif(TobArt, TobSuggest, TobCat : TOB);
procedure InitLigne(TOBPiece, TOBA, TOBL : TOB );
// Modif BTP
procedure decodeParametre (Params : string);
function  LanceTraitementFromStock (var TOBPiece : TOB) : boolean;
function  TraitDocuments (TOBpiece : TOB) : boolean;
function  RechercheDocuments (TOBLIsteDocuments : TOB) : boolean;
function  LanceTraitementFromDocs (var TOBPiece : TOB) : boolean;
procedure CalculeThisDocument(TOBPiece : TOB);
procedure DefiniListeDocuments (TOBListeDocuments : TOB ; All : boolean=false);
function  TraiteUneLigneBesoin (TobPiece,TOBL : TOB; LigneStandard : boolean ; var Ligne : integer) : boolean;
procedure EnleveLesExceptions(TobLig : tob);
procedure EnleveLesStandards(TOBLIG : TOB);
function  ConstitueRequeteDoc (sWhere : string; Detail : boolean=false) : string;
procedure FaisLeTriStd (TOBEntree : TOB);
procedure EnlevelesAutresExceptions (TOBLTraite,TOBLIG : TOB);

//procedure GenereBesoin (TOBART: TOB ;var TOBEvol : TOB;LigneStd : boolean ;TOBLTraite : TOB);
// --

implementation

uses
  ResulSuggestion,
  FactTarifs,
  UtilBTPgestChantier
  ;

var
    CleDoc      : R_CleDoc ;
    // Modif BTP
    TobArticles,TOBListeDocuments : TOB;
    TOBDetailStd,TOBDetailException : TOB;
    // ListNaturepiece : Tlist;
    TOBLIen : TOB;
    //
    AuMoinsUn   : Boolean;
    //TousDepots  : Boolean;
    //
    Depart, Valeur, Fermes, Nom, Histo, NbMois, ArtDeb, ArtFin, DepDeb, DepFin : string;
    Fam1, Fam2, Fam3, Fam1_, Fam2_, Fam3_, DatDeb, DatFin : string;
    Prio1, Prio2, Prio3, Prio4 : string;
    sPrefixe : string;
    NumLig : integer;
    // Modif BTP
    DepartTraitement : string;
    AvecHistoriqueConso,AvecSelectionDocument,WithStocks : boolean;
    //--



procedure DetermineQteAffecte (Lien : Integer; StkPhy : double);
var TOBD : TOB;
		reste : double;
begin
	reste := StkPhy;
  TOBD := TOBLien.FindFirst(['LIENLIGNE'],[Lien],true);
  repeat
  	if TOBD = nil then break;
    Reste := Reste - TOBD.GetValue('QTE');
    if Reste >= 0 then
    begin
    	TOBD.free;
    end else
    begin
    	TOBD.PutValue('BDA_QUANTITE',Reste * -1);
      Reste := 0;
    end;
  	TOBD := TOBLien.Findnext(['LIENLIGNE'],[Lien],true);
  until TOBD = nil;
end;

procedure decodeParametre (Params : string);
var select : string;
begin
Select := Params;
Depart := Trim(ReadTokenSt(Select));
Valeur := Trim(ReadTokenSt(Select));
Fermes := Trim(ReadTokenSt(Select));
Nom    := Trim(ReadTokenSt(Select));
Histo  := Trim(ReadTokenSt(Select));
NbMois := Trim(ReadTokenSt(Select));
ArtDeb := Trim(ReadTokenSt(Select));
ArtFin := Trim(ReadTokenSt(Select));
DepDeb := Trim(ReadTokenSt(Select));
DepFin := Trim(ReadTokenSt(Select));
Fam1   := Trim(ReadTokenSt(Select));
Fam1_  := Trim(ReadTokenSt(Select));
Fam2   := Trim(ReadTokenSt(Select));
Fam2_  := Trim(ReadTokenSt(Select));
Fam3   := Trim(ReadTokenSt(Select));
Fam3_  := Trim(ReadTokenSt(Select));
DatDeb := Trim(ReadTokenSt(Select));
DatFin := Trim(ReadTokenSt(Select));
Prio1  := Trim(ReadTokenSt(Select));
Prio2  := Trim(ReadTokenSt(Select));
Prio3  := Trim(ReadTokenSt(Select));
Prio4  := Trim(ReadTokenSt(Select));
// Modif BTP (Pour prise en compte des anciennes suggestion de réappro)
if Select <> '' then DepartTraitement := Trim(ReadTokenSt(Select)) else DepartTraitement := 'A';
if Select <> '' then AvecHistoriqueConso := (Trim(ReadTokenSt(Select))='H')
                else BEGIN
                     if (DatDeb <> DatFin) or (DatDeb <> DateToStr(Date()))
                      then
                        AvecHistoriqueConso := false
                      else
                        AvecHistoriqueConso := true;
                     END;
if Select <> '' then AvecSelectionDocument := (Trim(ReadTokenSt(Select))='D')
                else AvecSelectionDocument := false;

if (Select <> '') then WithStocks := (Trim(ReadTokenSt(Select))='O')
									else WithStocks := True;
end;

// MODIFS LS
function IsPieceCC (Naturepiece : string) : boolean;
var UneNature,LesNatureCC : string;
//    indice : integer;
begin
  result := false;
  lesNatureCC := GetNaturePieceCde (false);
  UneNature := READTOKENST (lesNatureCC);
  repeat
    if UneNature = '' then break;
    if UneNature = Naturepiece then
    begin
      result := true;
      break;
    end;
    UneNature := READTOKENST (lesNatureCC);
  until UneNature = '';
end;

function IsPieceBLC (Naturepiece : string) : boolean;
var UneNature,LesNatureBLC : string;
//    indice : integer;
begin
  result := false;
  lesNatureBLC := GetNaturePieceBLC (false);
  UneNature := READTOKENST (lesNatureBLC);
  repeat
    if UneNature = '' then break;
    if UneNature = Naturepiece then
    begin
      result := true;
      break;
    end;
    UneNature := READTOKENST (lesNatureBLC);
  until UneNature = '';
end;

function IsPieceCForBLForFF  (Naturepiece : string) : boolean;
var {UneNature,}LesNatureCF : string;
//    indice : integer;
begin
//  result := false;
  lesNatureCF := GetPieceAchat (true,true);
{
  UneNature := READTOKENST (lesNatureCF);
  repeat
    if UneNature = '' then break;
    if UneNature = Naturepiece then
    begin
      result := true;
      break;
    end;
    UneNature := READTOKENST (lesNatureCF);
  until UneNature = '';
}
	Result := (Pos(NaturePiece,LesNatureCf)>0);
end;

function IsReferencePresente (TheMode : IrpMode; TobRef : TOB): boolean;
var TheClef : R_Cledoc;
    TheResult : TOB;
begin
  result := false;
  if TheMode = IrpPrec then
  begin
    if TOBRef.GetValue('GL_PIECEPRECEDENTE') = '' then BEGIN Exit; END;
    DecodeRefPiece (TOBRef.GetValue('GL_PIECEPRECEDENTE'),TheClef);
    TheResult := TOBListeDocuments.FindFirst (['GP_NATUREPIECEG','GP_NUMERO','GP_INDICEG'],
                                     [TheClef.Naturepiece,TheClef.numeroPiece,
                                     TheClef.indice],true);
    if TheResult <> nil then result := true;
  end else if TheMode = IrpOrig then
  begin
    if TOBRef.GetValue('GL_PIECEORIGINE') = '' then BEGIN Exit; END;
    DecodeRefPiece (TOBRef.GetValue('GL_PIECEORIGINE'),TheClef);
    TheResult := TOBListeDocuments.FindFirst (['GP_NATUREPIECEG','GP_NUMERO','GP_INDICEG'],
                                     [TheClef.Naturepiece,TheClef.numeroPiece,
                                     TheClef.indice],true);
    if TheResult <> nil then result := true;
  end else if TheMode = IrpRef then
  begin
    TheResult := TOBListeDocuments.FindFirst (['GP_NATUREPIECEG','GP_NUMERO','GP_INDICEG'],
                                     [TobRef.GetValue('GL_NATUREPIECEG'),TobRef.GetValue('GL_NUMERO'),
                                     TobRef.GetValue('GL_INDICEG')],true);
    if TheResult <> nil then result := true;
  end;
end;

function IsPieceOK (TOBL : TOB) : boolean;
var VenteAchat,Naturepiece : string;
begin
  if TOBListeDocuments = nil then BEGIN Result := True; Exit; END;
  if TOBListeDocuments.detail.count = 0 then BEGIN Result := True; Exit; END;
  result := false;
  Naturepiece := TOBL.GetValue('GL_NATUREPIECEG');
  VenteAchat := GetInfoParPiece(Naturepiece, 'GPP_VENTEACHAT');
  // Piece de commande client
  if (IsPieceCC(Naturepiece)) and (ISReferencePresente (IrpRef,TOBL)) then
  begin
    result := true;
    exit;
  END else if (IsPieceBLC(NaturePiece)) and
     (IsReferencePresente (IrpPrec,TOBL)) then
  BEGIN
    result := true;
    exit;
  END else if (IsPieceCForBLForFF (Naturepiece)) and
          (IsReferencePresente (IrpOrig,TOBL)) then
  BEGIN
    result := true;
    exit;
  END;

end;
// -----

function LanceTraitementFromStock (var TOBPiece : TOB) : boolean;
var TobTemp : TOB;
    s_where : String;
    Select  : string;
    Q       : TQuery;
    i_ind1  : integer;
begin

  result := false;

  //FV1 - 30/11/2015 - FS#1772 - SOPREL : la réappro pour besoin de stocks ne fonctionne pas
  //if (DepDeb = '') and (DepFin = '') then TousDepots := True else TousDepots := False;
  Select := 'Select * from ARTICLE where ';

  if (ArtDeb <> '') and (ArtFin <> '') then
      Select := Select + 'GA_CODEARTICLE>="' + ArtDeb + '" AND GA_CODEARTICLE<="' + ArtFin + '" '
  else if ArtDeb <> '' then
      Select := Select + 'GA_CODEARTICLE>="' + ArtDeb + '" '
  else if ArtFin <> '' then
      Select := Select + 'GA_CODEARTICLE<="' + ArtFin + '" ';

  if (ArtDeb <> '') or (ArtFin <> '') then Select := Select + ' AND ';
  Select := Select + 'GA_TENUESTOCK="X" and GA_STATUTART<>"GEN" and ';
  if Fermes = 'N' then
    Select := Select + '(GA_FERME<>"X") '
  else if Fermes = 'O' then
    Select := Select + '(GA_FERME="X") '
  else
  begin
    Select := Select + '((GA_FERME<>"X") or (GA_FERME="X" and ';
    Select := Select + 'GA_DATESUPPRESSION>"' + USDateTime(StrToDate(DatDeb)) + '")) ';
  end;

  if Fam1 <> '' then
  begin
    if Fam1_ <> '' then
      Select := Select + 'and (GA_FAMILLENIV1>="' + Fam1 + '" and GA_FAMILLENIV1<="' + Fam1_ + '") '
    else
      Select := Select + 'and (GA_FAMILLENIV1>="' + Fam1 + '" ';
  end;

  if Fam2 <> '' then
  begin
    if Fam2_ <> '' then
      Select := Select + 'and (GA_FAMILLENIV2>="' + Fam2 + '" and GA_FAMILLENIV2<="' + Fam2_ + '") '
    else
      Select := Select + 'and GA_FAMILLENIV2="' + Fam2 + '" ';
  end;

  if Fam3 <> '' then
  begin
    if Fam3_ <> '' then
      Select := Select + 'and (GA_FAMILLENIV3>="' + Fam3 + '" and GA_FAMILLENIV3<="' + Fam3_ + '") '
    else
      Select := Select + 'and GA_FAMILLENIV3="' + Fam3 + '" ';
  end;

  s_where := '';

  Q := OpenSQL(Select, True,-1,'',true);

  if Q.EOF then
  begin
    Ferme(Q);
    Exit;
  end;

  TobArticles := TOB.Create('LES ARTICLES', nil, -1);
  TobArticles.LoadDetailDB('ARTICLE', '', '', Q, False);
  Ferme(Q);
  s_where := '';
  InitMove(TobArticles.Detail.Count, '');
  NumLig := 0;
  for i_ind1 := 0 to TobArticles.Detail.Count - 1 do
  begin
    TobTemp := TobArticles.Detail[i_ind1];
    TraiteUnArticle(TOBPiece, TobTemp);
    MoveCur(False);
  end;

  FiniMove();

  CalculeThisDocument(TOBPiece);

  FResulSuggestion.GZZ_SUGGEST.Text := TOBPiece.GetValue('GP_NUMERO');

  Result := (AuMoinsUn) and (tobpiece.detail.count > 0);

  TobArticles.Free;

end;

function Calcul_Suggestion(Params : string; var TobPiece : TOB; var TOBLDEVCHAR : TOB) : boolean;
begin
  //
  TOBListeDocuments   := nil;
  TOBDetailException  := nil;
  TOBDetailStd        := nil;
  TOBLien             := TOBLDEVCHAR;

  //  Création du select des articles en fonction des paramètres saisis
  AuMoinsUn := False;
  decodeParametre (Params);

  // Definition du TOBPiece (initialisation)
  sPrefixe            := TableToPrefixe('GCTMPREAPPRO');
  CleDoc.NaturePiece  :='REA' ;
  CleDoc.DatePiece    :=V_PGI.DateEntree ;
  CleDoc.Souche       :='' ;
  CleDoc.NumeroPiece  :=0 ;
  CleDoc.Indice       :=0 ;
  CleDoc.Souche       :=GetSoucheG(CleDoc.NaturePiece,VH^.EtablisDefaut,'') ;
  CleDoc.NumeroPiece  :=GetNumSoucheG(CleDoc.Souche,CleDoc.DatePiece) ;

  InitTOBPiece(TobPiece);

  TOBPiece.PutValue('GP_NUMERO',CleDoc.NumeroPiece) ;
  TOBPiece.PutValue('GP_SOUCHE',CleDoc.Souche) ;
  TOBPiece.PutValue('GP_INDICEG',CleDoc.Indice) ;
  TOBPiece.PutValue('GP_CREEPAR','MAN') ;
  TOBPiece.PutValue('GP_SOCIETE',V_PGI.CodeSociete);
  TOBPiece.PutValue('GP_ETABLISSEMENT',VH^.EtablisDefaut);
  TOBPiece.PutValue('GP_DATEPIECE',V_PGI.DateEntree);
  TOBPiece.PutValue('GP_DATECREATION',V_PGI.DateEntree);
  TOBPiece.PutValue('GP_DATEMODIF',V_PGI.DateEntree);
  TOBPiece.PutValue('GP_DEVISE',V_PGI.DevisePivot);
  TOBPiece.PutValue('GP_FACTUREHT', 'X');
  TOBPiece.PutValue('GP_NATUREPIECEG','REA') ;
  TOBPiece.PutValue('GP_VENTEACHAT',GetInfoParPiece('REA','GPP_VENTEACHAT')) ;
  TOBPiece.PutValue('GP_DEPOT', '');

  if DepartTraitement = 'A' then
    Result := LanceTraitementFromStock (TOBPiece)
  else
    Result := LanceTraitementFromDocs (TOBPiece);
  // --
end;

procedure TraiteUnArticle(var TobPiece : TOB; TOBArt : Tob);
var TobMvts : Tob;
    TobEvol : TOB;
    st1     : string;
begin
  // Modif BTP
  {$IFDEF BTP}
  if IsPrestationInterne (TOBART.GetValue('GA_ARTICLE')) then Exit;
  {$ENDIF}

  TobMvts := TOB.Create('', nil, -1);
  TobEvol := TOB.Create('', nil, -1);
  st1 := VH_GC.GCPeriodeStock;

  (* if (DatDeb = DatFin) and (DatDeb = DateToStr(Date())) then *)
  if (not AvecHistoriqueConso) and (TOBArt.GetValue('GA_TENUESTOCK')='X') then
  begin
    CalculRupture(TobPiece, TOBArt, TobMvts, TobEvol);
    if TobEvol.Detail.Count > 0 then EnregDonnees(TobPiece, TOBArt, TobEvol, TobMvts,nil);
  end
  else
  begin
    CalculHistoStock(TOBArt, TobMvts);
    CalculStockPhysique(TOBArt, TobEvol);
    CalculEvolutionStock(TOBArt, TobMvts, TobEvol);
    // Modif BTP
    if (TobEvol.Detail.Count > 0) or (TobMvts.Detail.Count > 0) then
    // --
    EnregDonnees(TobPiece, TOBArt, TobEvol, TobMvts,nil);
  end;

  if TobEvol.Detail.Count > 0 then AuMoinsUn := True;

  TobMvts.PutGridDetail(FResulSuggestion.G_Mvts,True,True,'PERIODE;DEPOT;QENT;QSOR;COEFENT;COEFSOR');
  TobEvol.PutGridDetail(FResulSuggestion.G_Evol,True,True,'DEPOT;DATE;PIECE;Q_PHYSIQUE;Q_CLIENT;Q_FOURNI');
  //TobPiece.PutGridDetail(FResulSuggestion.G_Fin,True,True,'GL_CODEARTICLE;GL_LIBELLE;GL_DEPOT;GL_QTESTOCK;GL_PUHT;GL_TIERS');

  TobMvts.Free;
  TobEvol.Free;

  TobMvts := nil;
  TobEvol := nil;

end;

procedure CalculHistoStock(TOBArt : TOB; var TobMvts : TOB; LigneStd : boolean);
var TobLig, TobTempL, TobTempM : TOB;
    i_ind1 : integer;
    Q : TQuery;
    st1, st2, Mois, Annee, Depot, Select : string;
begin
  // Dans le cas ou l'on traite une ligne "specifique" on a pas besoin de calculer l'histo des mvts
  if not LigneStd then Exit;
  if (not AvecHistoriqueConso) or (TOBART.GetValue('GA_TENUESTOCK')='-') or (not WithStocks) then exit;
  //  Construction des entrees - sorties mois par mois
  Select:='Select GQ_DATECLOTURE,GQ_ARTICLE,GQ_DEPOT,GQ_CUMULSORTIES,GQ_CUMULENTREES' +
          ' from DISPO ' +
          'where GQ_ARTICLE="' + TOBArt.GetValue('GA_ARTICLE') + '" and ';
  if Histo = 'L' then
    Select := Select + 'GQ_DATECLOTURE>="' + USDateTime(Date - 731) + '"'
  else
    Select := Select + 'GQ_DATECLOTURE>="' + USDateTime(Plusmois(Date, -StrToInt(NbMois))) + '"';
  Select := Select + ' and GQ_DATECLOTURE<="' + USDateTime(Date) + '" ';
  Q := OpenSQL(Select, True,-1,'',true);
  if Q.EOF then
  begin
    Ferme(Q);
    Exit;
  end;
  //
  TobLig := TOB.Create('', nil, -1);
  TobLig.LoadDetailDB('DISPO', '', '', Q, False);
  Ferme(Q);
  for i_ind1 := 0 to TobLig.Detail.count - 1 do
  begin
    TobTempL := TobLig.Detail[i_ind1];
    st1 := DateToStr(TobTempL.GetValue('GQ_DATECLOTURE'));
    st2 := ReadTokenPipe(st1, '/');
    Mois  := ReadTokenPipe(st1, '/');
    Annee := ReadTokenPipe(st1, '/');
    if VH_GC.GCPeriodeStock = 'QUI' then
        if StrToInt(st2) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
    else if VH_GC.GCPeriodeStock = 'SEM' then
        Mois := Format('%.2d', [NumSemaine(TobTempL.GetValue('GQ_DATECLOTURE'))]);

    //FV1 - 30/11/2015 - FS#1772 - SOPREL : la réappro pour besoin de stocks ne fonctionne pas
    //if TousDepots then Depot := '' else
    Depot := TobTempL.GetValue('GQ_DEPOT');
    TobTempM := CreerFilleMvts(TobMvts, Depot, Annee, Mois);
    //
    TobTempM.PutValue('QENT', TobTempM.GetValue('QENT') + TobTempL.GetValue('GQ_CUMULENTREES'));
    TobTempM.PutValue('QSOR', TobTempM.GetValue('QSOR') + TobTempL.GetValue('GQ_CUMULSORTIES'));
  end;

  //Calcul des coef pondérateurs
  for i_ind1 := 0 to TobMvts.Detail.Count - 1 do
  begin
    TobTempL := TobMvts.Detail[i_ind1];
    st1 := Copy(TobTempL.GetValue('PERIODE'), 0, 4);
    st2 := IntToStr(StrToInt(st1) - 1) + Copy(TobTempL.GetValue('PERIODE'), Pos('/', TobTempL.GetValue('PERIODE')), 255);
    //FV1 - 30/11/2015 - FS#1772 - SOPREL : la réappro pour besoin de stocks ne fonctionne pas
    //if not TousDepots then st1 := TobTempL.GetValue('DEPOT');
    st1 := TobTempL.GetValue('DEPOT');
    TobTempM := TobMvts.FindFirst(['PERIODE','DEPOT'], [st2, st1], False);
    if TobTempM <> nil then
    begin
      if TobTempM.GetValue('QENT') <> 0 then
          TobTempL.PutValue('COEFENT', TobTempL.GetValue('QENT') / TobTempM.GetValue('QENT'));
      if TobTempM.GetValue('QSOR') <> 0 then
          TobTempL.PutValue('COEFSOR', TobTempL.GetValue('QSOR') / TobTempM.GetValue('QSOR'));
    end;
  end;
  TobLig.Free;
end;

procedure CalculStockPhysique(TOBArt : TOB; var TobEvol : TOB);
var
    TobTempE, TobTemp, TobDispo : TOB;
    Q : TQuery;
    i_ind1 : integer;
    Select, Annee, Mois, Jour : string;

begin
  // MOdif BTP
  if (TOBART.GetValue('GA_TENUESTOCK')='-') or (not WithStocks) then exit;
  // --
  //  Calcul du stock physique, tous depot ou en fonction des depot choisis
  TobDispo := TOB.Create('', nil, -1);
  TRY
    Select := 'Select GQ_DEPOT,GQ_PHYSIQUE from DISPO where GQ_ARTICLE="' +
              TOBArt.GetValue('GA_ARTICLE') + '" and GQ_CLOTURE="-" ';
    if DepDeb <> '' then Select := Select + 'and GQ_DEPOT>="' + DepDeb + '" ';
    if DepFin <> '' then Select := Select + 'and GQ_DEPOT<="' + DepFin + '" ';
    Q := OpenSQL(Select, True,-1,'',true);
    if not Q.EOF then
    begin
      TobDispo.LoadDetailDB('DISPO', '', '', Q, False);
      {*
      //FV1 - 30/11/2015 - FS#1772 - SOPREL : la réappro pour besoin de stocks ne fonctionne pas
      if TousDepots then
      begin
        EclateDate(Date, Annee, Mois, Jour);
        if VH_GC.GCPeriodeStock = 'QUI' then
            if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
        else if VH_GC.GCPeriodeStock = 'SEM' then
            Mois := Format('%.2d', [NumSemaine(TobDispo.Detail[0].GetValue('GQ_DATECLOTURE'))]);
        TobTempE := CreerFilleEvol(TobEvol, '', Annee, Mois);
        for i_ind1 := 0 to TobDispo.Detail.Count - 1 do
            TobTempE.PutValue('Q_PHYSIQUE', TobTempE.GetValue('Q_PHYSIQUE') +
                                            TobDispo.Detail[i_ind1].GetValue('GQ_PHYSIQUE'));
        GenereEvolPeriode(TobEvol, '');
      end
      else
      begin
      *}
      for i_ind1 := 0 to TobDispo.Detail.Count - 1 do
      begin
        TobTemp := TobDispo.Detail[i_ind1];
        EclateDate(Date, Annee, Mois, Jour);
        if VH_GC.GCPeriodeStock = 'QUI' then
            if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
        else if VH_GC.GCPeriodeStock = 'SEM' then
          Mois := Format('%.2d', [NumSemaine(TobTemp.GetValue('GQ_DATECLOTURE'))]);
        TobTempE := CreerFilleEvol(TobEvol, TobTemp.GetValue('GQ_DEPOT'), Annee, Mois);
        TobTempE.PutValue('Q_PHYSIQUE', TobTempE.GetValue('Q_PHYSIQUE') +
                                        TobTemp.GetValue('GQ_PHYSIQUE'));
        GenereEvolPeriode(TobEvol, TobTemp.GetValue('GQ_DEPOT'));
      end;
      //end;
    end;
    Ferme(Q);
  FINALLY
    TobDispo.free;
  END;
end;

(*
procedure GenereBesoin (TOBART: TOB; var TOBEvol : TOB;LigneStd : boolean ;TOBLTraite : TOB);
var
    TobTempE, TobTempL, TobLig : TOB;
    Q : TQuery;
    stkphy, FUS, FUA, FUV : double;
    i_ind1, i_ind2 : integer;
    Select, st1, st2, Annee, Mois, Jour : string;
    NewPeriode : boolean;
    d_temp1 : TDateTime;
begin
//  selection des lignes de date à venir
Select:='Select GL_NATUREPIECEG,GL_NUMERO,GL_SOUCHE,GL_INDICEG,GL_NUMLIGNE,GL_DATEPIECE,GL_ARTICLE,GL_DEPOT,GL_QTESTOCK,GL_QTEFACT,' +
        'GL_QTERESTE,GL_DATELIVRAISON,GL_QUALIFQTESTO,GL_QUALIFQTEVTE,GL_QUALIFQTEACH,GP_VIVANTE from LIGNE ' +
        'left join PARPIECE on GPP_NATUREPIECEG=GL_NATUREPIECEG ' +
        'left join PIECE on (GP_NATUREPIECEG=GL_NATUREPIECEG) and (GP_NUMERO=GL_NUMERO) ' +
        'where GL_ARTICLE="' + TOBArt.GetValue('GA_ARTICLE') + '" ' +
        'and GP_VIVANTE="X" ' +
        'and GL_VIVANTE="X" ' +
        'and ((GL_DATELIVRAISON>"' + USDateTime(Date) + '") ' +
//        'and ((GL_DATELIVRAISON>="' + USDateTime(StrToDate(DatDeb)) + '" ' +
        'or (GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' +
        'and GPP_VENTEACHAT="VEN") ' +
        'or (GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' +
        'and GP_ETATVISA<>"ATT" ' +
        'and GPP_VENTEACHAT="ACH"))';
Q := OpenSQL(Select, True);
if Q.EOF then
    begin
    Ferme(Q);
    Exit;
    end;
TobLig := TOB.Create('', nil, -1);
TobLig.LoadDetailDB('LIGNE', '', '', Q, False);
Ferme(Q);
// Modif BTP pour ne traiter que le cas de figure qui se présente
if LigneStd then
BEGIN
  EnleveLesExceptions(TobLig)
END ELSE
BEGIN
  EnleveLesStandards(TOBLIG);
  EnlevelesAutresExceptions (TOBLTraite,TOBLIG);
END;
// ----
//  calcul de l'évolution du stock dans le futur. sur les lignes selectionnées, on va
//  trier les cdes clients et fournisseurs.
for i_ind1 := 0 to TobLig.Detail.Count - 1 do
    begin
    TobTempL := TobLig.Detail[i_ind1];

    FUS := RatioMesure('PIE', TobTempL.GetValue('GL_QUALIFQTESTO'));
    FUV := RatioMesure('PIE', TobTempL.GetValue('GL_QUALIFQTEVTE'));
    FUA := RatioMesure('PIE', TobTempL.GetValue('GL_QUALIFQTEACH'));

    st1 := GetInfoParPiece(TobTempL.GetValue('GL_NATUREPIECEG'),'GPP_QTEPLUS');
    st2 := GetInfoParPiece(TobTempL.GetValue('GL_NATUREPIECEG'),'GPP_QTEMOINS');
    i_ind2 := Pos('RC', st1) + Pos('RC', st2) + Pos('RF', st1) + Pos('RF', st2)+ Pos('PRE',st1) + Pos('PRE',st2);
    if i_ind2 = 0 then Continue;
    if (Pos('RC', st1) + Pos('RC', st2) + Pos('PRE',st1) + Pos('PRE',st2)) <> 0 then
        begin
        d_temp1 := TobTempL.GetValue('GL_DATEPIECE');
        TobTempL.PutValue('GL_QTESTOCK', (TobTempL.GetValue('GL_QTERESTE') * FUV) / FUS);
        end;
    if (Pos('RF', st1) + Pos('RF', st2) <> 0) then
        begin
        d_temp1 := TobTempL.GetValue('GL_DATELIVRAISON');
        TobTempL.PutValue('GL_QTESTOCK', (TobTempL.GetValue('GL_QTERESTE') * FUA) / FUS);
        end;
    EclateDate(d_temp1, Annee, Mois, Jour);
    NewPeriode := True;
    if VH_GC.GCPeriodeStock = 'QUI' then
        if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
    else if VH_GC.GCPeriodeStock = 'SEM' then
        Mois := Format('%.2d', [NumSemaine(d_temp1)]);
    EclateDate(Date, Annee, Mois, Jour);
    if VH_GC.GCPeriodeStock = 'QUI' then
        if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
    else if VH_GC.GCPeriodeStock = 'SEM' then
        Mois := Format('%.2d', [NumSemaine(TobTempL.GetValue('GL_DATEPIECE'))]);
    EclateDate(d_temp1, Annee, Mois, Jour);
    if VH_GC.GCPeriodeStock = 'QUI' then
        if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
    else if VH_GC.GCPeriodeStock = 'SEM' then
        Mois := Format('%.2d', [NumSemaine(d_temp1)]);
    stkphy := 0.0;
    if NewPeriode then
        if TousDepots then
            TobTempE := CreerFilleEvol(TobEvol, '', Annee, Mois)
            else
            TobTempE := CreerFilleEvol(TobEvol, TobTempL.GetValue('GL_DEPOT'), Annee, Mois);
    if not TousDepots then TobTempE.PutValue('DEPOT', TobTempL.GetValue('GL_DEPOT'));
    if (Pos('RC', st1)+ Pos('PRE',st1)) <> 0  then
        TobTempE.PutValue('Q_CLIENT', TobTempE.GetValue('Q_CLIENT') + TobTempL.GetValue('GL_QTERESTE'));
    if (Pos('RC', st2)+ Pos('PRE',st2)) <> 0 then
        TobTempE.PutValue('Q_CLIENT', TobTempE.GetValue('Q_CLIENT') + TobTempL.GetValue('GL_QTERESTE') * -1);
    if Pos('RF', st1) <> 0 then
        TobTempE.PutValue('Q_FOURNI', TobTempE.GetValue('Q_FOURNI') + TobTempL.GetValue('GL_QTERESTE'));
    if Pos('RF', st2) <> 0 then
        TobTempE.PutValue('Q_FOURNI', TobTempE.GetValue('Q_FOURNI') + TobTempL.GetValue('GL_QTERESTE') * -1);
    end;
TobLig.Free;
end;
*)
procedure CalculEvolutionStock(TOBArt, TobMvts : TOB; var TobEvol : TOB;LigneStd:boolean=true;TOBLTraite : TOB = nil;FromCalculDocuments : boolean=false; Libart : String='');
var
    TobTempE, TobTempL, TobLig : TOB;
    Q : TQuery;
    FUS, FUA, FUV, CoefUaUs,CoefUsUv : double;
    i_ind1, i_ind2 : integer;
    Select, st1, st2, Annee, Mois, Jour : string;
    NewPeriode : boolean;
    d_temp1 : TDateTime;

begin
  // MODIF BTP
  if (not AvecHistoriqueConso) AND (TOBART.GEtValue('GA_TENUESTOCK')='X') then Exit;
  // if (DatDeb = DatFin) and (DatDeb = DateToStr(Date())) then Exit;
  // --
  //  selection des lignes de date à venir
  Select:='Select GL_NATUREPIECEG,GL_NUMERO,GL_SOUCHE,GL_INDICEG,GL_NUMLIGNE,GL_DATEPIECE,GL_ARTICLE,' +
        (* Ajout Ls *)
        'GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_COEFCONVQTE,GL_COEFCONVQTEVTE,' +
        (* --- *)
          'GL_DEPOT,GL_QTESTOCK,GL_QTEFACT, GL_QTERESTE, GL_MTRESTE, GL_DATELIVRAISON,GL_QUALIFQTESTO,GL_QUALIFQTEVTE,GL_QUALIFQTEACH,GP_VIVANTE from LIGNE ' +
        'left join PARPIECE on GPP_NATUREPIECEG=GL_NATUREPIECEG ' +
        'left join PIECE on (GP_NATUREPIECEG=GL_NATUREPIECEG) and (GP_NUMERO=GL_NUMERO) ' +
        'where GL_ARTICLE="' + TOBArt.GetValue('GA_ARTICLE') + '" ' +
        'and GL_LIBELLE="' + Libart + '" ' +
        'and GP_VIVANTE="X" ' +
        'and GL_VIVANTE="X" ' ;

  if Assigned(TOBLTraite) then
  begin
	if TOBLTraite.getValue('GL_AFFAIRE')<>'' then
  begin
		Select := Select + 'AND GL_AFFAIRE="'+TOBLTraite.getValue('GL_AFFAIRE')+'" '
  end;
  end;

if DepDeb  <> '' then
  if DepFin  <> '' then
    Select := Select + 'AND (GL_DEPOT>="' + DepDeb + '" AND GL_DEPOT<="' + DepFin + '") '
  else
    Select := Select + 'AND GL_DEPOT="' + DepDeb + '" ';

  //        'and ((GL_DATELIVRAISON>"' + USDateTime(Date) + '") ' +
  Select := Select + 'and ((GL_DATELIVRAISON>="' + USDateTime(StrToDate(DatDeb)) + '") ' +
                   'or (GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' +
                   'and GPP_VENTEACHAT="VEN") ' +
                   'or (GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' +
                   'and GP_ETATVISA<>"ATT" ' +
                   'and GPP_VENTEACHAT="ACH"))';
  Q := OpenSQL(Select, True,-1,'',true);

  if Q.EOF then
    begin
    Ferme(Q);
    Exit;
    end;

  TobLig := TOB.Create('', nil, -1);
  TobLig.LoadDetailDB('LIGNE', '', '', Q, False);

  Ferme(Q);

  Select:='Select GL_NATUREPIECEG,GL_NUMERO,GL_SOUCHE,GL_INDICEG,GL_NUMLIGNE,GL_DATEPIECE,GL_ARTICLE,' +
        (* Ajout Ls *)
        'GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_COEFCONVQTE,GL_COEFCONVQTEVTE,' +
        (* --- *)
        'GL_DEPOT,GL_QTESTOCK,GL_QTEFACT,GL_QTERESTE, GL_MTRESTE, ' +
        'GL_DATELIVRAISON,GL_QUALIFQTESTO,GL_QUALIFQTEVTE,GL_QUALIFQTEACH, ' +
        '(GL_QTERESTE * GLN_QTE) as GL_QTERESTE, GL_MTRESTE, GL_DATELIVRAISON,GP_VIVANTE from LIGNENOMEN ' +
        'left join LIGNE on (GLN_NATUREPIECEG=GL_NATUREPIECEG) and (GLN_NUMERO=GL_NUMERO) ' +
        'and (GLN_NUMLIGNE=GL_NUMLIGNE) and (GLN_INDICEG=GL_INDICEG) ' +
        'left join PIECE on (GP_NATUREPIECEG=GL_NATUREPIECEG) and (GP_NUMERO=GL_NUMERO) ' +
        'left join PARPIECE on GPP_NATUREPIECEG=GL_NATUREPIECEG ' +
        'where GLN_ARTICLE="' + TOBArt.GetValue('GA_ARTICLE') + '" ' +
        'and GLN_LIBELLE="' + Libart + '" ' +
        'and GP_VIVANTE="X" ' +
        'and GL_VIVANTE="X" ' ;
if DepDeb  <> '' then
  if DepFin  <> '' then
    Select := Select + 'AND (GL_DEPOT>="' + DepDeb + '" AND GL_DEPOT<="' + DepFin + '") '
  else
    Select := Select + 'AND GL_DEPOT="' + DepDeb + '" ';
//        'and ((GL_DATELIVRAISON>"' + USDateTime(Date) + '") ' +
Select := Select + 'and ((GL_DATELIVRAISON>="' + USDateTime(StrToDate(DatDeb)) + '") ' +
                   'or (GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' +
                   'and GPP_VENTEACHAT="VEN") ' +
                   'or (GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' +
                   'and GP_ETATVISA<>"ATT" ' +
                   'and GPP_VENTEACHAT="ACH"))';
Q := OpenSQL(Select, True,-1,'',true);
if Q.EOF then
    begin
    Ferme(Q);
    end
    else
    begin
//    TobLig := TOB.Create('', nil, -1);
    TobLig.LoadDetailDB('LIGNE', '', '', Q, True);
    Ferme(Q);
    end;
// Modif BTP pour ne traiter que le cas de figure qui se présente

if LigneStd then
BEGIN
  EnleveLesExceptions(TobLig)
END ELSE
BEGIN
  EnleveLesStandards(TOBLIG);
  EnlevelesAutresExceptions (TOBLTraite,TOBLIG);
END;
// ----
//  calcul de l'évolution du stock dans le futur. sur les lignes selectionnées, on va
//  trier les cdes clients et fournisseurs.
for i_ind1 := 0 to TobLig.Detail.Count - 1 do
    begin
    TobTempL := TobLig.Detail[i_ind1];
    if not IsPieceOK (TobTempL) then continue;
    COEFUAUS := TobTempL.getValue('GL_COEFCONVQTE');
    CoefUsUv := TobTempL.getValue('GL_COEFCONVQTEVTE');
    FUS := RatioMesure('PIE', TobTempL.GetValue('GL_QUALIFQTESTO')); // ??
    FUV := RatioMesure('PIE', TobTempL.GetValue('GL_QUALIFQTEVTE'));
    FUA := RatioMesure('PIE', TobTempL.GetValue('GL_QUALIFQTEACH'));

    st1 := GetInfoParPiece(TobTempL.GetValue('GL_NATUREPIECEG'),'GPP_QTEPLUS');
    st2 := GetInfoParPiece(TobTempL.GetValue('GL_NATUREPIECEG'),'GPP_QTEMOINS');
    i_ind2 := Pos('RC', st1) + Pos('RC', st2) + Pos('RF', st1) + Pos('RF', st2)+ Pos('PRE',st1) + Pos('PRE',st2);
    if i_ind2 = 0 then Continue;
    d_temp1 := TobTempL.GetValue('GL_DATEPIECE');
    if (Pos('RC', st1) + Pos('RC', st2) + Pos('PRE',st1) + Pos('PRE',st2)) <> 0 then
        begin
        if CoefUsUv <> 0 then
        begin
        	TobTempL.PutValue('GL_QTERESTE', TobTempL.GetValue('GL_QTERESTE') / CoefUSUV);
        end else
        begin
        	TobTempL.PutValue('GL_QTERESTE', (TobTempL.GetValue('GL_QTERESTE') * FUV) / FUS);
        end;
        end;
    if (Pos('RF', st1) + Pos('RF', st2) <> 0) then
        begin
          d_temp1 := TobTempL.GetValue('GL_DATELIVRAISON');
          if COEFUAUS <> 0 then
          begin
          	TobTempL.PutValue('GL_QTERESTE', (TobTempL.GetValue('GL_QTERESTE') * CoefUaUs));
          end else
          begin
          	TobTempL.PutValue('GL_QTERESTE', (TobTempL.GetValue('GL_QTERESTE') * FUA) / FUS);
          end;
        end;
    EclateDate(d_temp1, Annee, Mois, Jour);
    NewPeriode := True;
    if VH_GC.GCPeriodeStock = 'QUI' then
        if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
    else if VH_GC.GCPeriodeStock = 'SEM' then
        Mois := Format('%.2d', [NumSemaine(d_temp1)]);

    //FV1 - 30/11/2015 - FS#1772 - SOPREL : la réappro pour besoin de stocks ne fonctionne pas
    //if TousDepots then
    //    TobTempE := TobEvol.FindFirst(['DEPOT','DATE'], ['', Annee + '/' + Mois], False)
    //    else
    TobTempE := TobEvol.FindFirst(['DEPOT','DATE'],[TobTempL.GetValue('GL_DEPOT'), Annee + '/' + Mois], False);
    if TobTempE <> nil then
        begin
        NewPeriode := False;
        end
        else
        begin
        EclateDate(Date, Annee, Mois, Jour);
        if VH_GC.GCPeriodeStock = 'QUI' then
            if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
        else if VH_GC.GCPeriodeStock = 'SEM' then
            Mois := Format('%.2d', [NumSemaine(TobTempL.GetValue('GL_DATEPIECE'))]);
        TobTempE := TobEvol.FindFirst(['DEPOT','DATE'], ['', Annee + '/' + Mois], False);
        EclateDate(d_temp1, Annee, Mois, Jour);
        if VH_GC.GCPeriodeStock = 'QUI' then
            if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
        else if VH_GC.GCPeriodeStock = 'SEM' then
            Mois := Format('%.2d', [NumSemaine(d_temp1)]);
        end;
//    if TobTempE <> nil then stkphy := TobTempE.GetValue('Q_PHYSIQUE') else stkphy := 0.0;
    if NewPeriode then
      //FV1 - 30/11/2015 - FS#1772 - SOPREL : la réappro pour besoin de stocks ne fonctionne pas
      //if TousDepots then
      //    TobTempE := CreerFilleEvol(TobEvol, '', Annee, Mois)
      //    else
      TobTempE := CreerFilleEvol(TobEvol, TobTempL.GetValue('GL_DEPOT'), Annee, Mois);
      //FV1 - 30/11/2015 - FS#1772 - SOPREL : la réappro pour besoin de stocks ne fonctionne pas
      //if not TousDepots then
    TobTempE.PutValue('DEPOT', TobTempL.GetValue('GL_DEPOT'));
    if (Pos('RC', st1)+ Pos('PRE',st1)) <> 0  then
        TobTempE.PutValue('Q_CLIENT', TobTempE.GetValue('Q_CLIENT') + TobTempL.GetValue('GL_QTERESTE'));
    if (Pos('RC', st2)+ Pos('PRE',st2)) <> 0 then
        TobTempE.PutValue('Q_CLIENT', TobTempE.GetValue('Q_CLIENT') + TobTempL.GetValue('GL_QTERESTE') * -1);
    if Pos('RF', st1) <> 0 then
        TobTempE.PutValue('Q_FOURNI', TobTempE.GetValue('Q_FOURNI') + TobTempL.GetValue('GL_QTERESTE'));
    if Pos('RF', st2) <> 0 then
        TobTempE.PutValue('Q_FOURNI', TobTempE.GetValue('Q_FOURNI') + TobTempL.GetValue('GL_QTERESTE') * -1);
    end;
TobLig.Free;
//  on compare les resultats obtenus avec l'historique des cours
for i_ind1 := 0 to TobEvol.Detail.Count - 1 do
    begin
    TobTempE := TobEvol.Detail[i_ind1];
    st2 := TobTempE.GetValue('DATE');
    st1 := IntToStr(StrToInt(ReadTokenPipe(st2,'/')) - 1);
    TobTempL := TobMvts.FindFirst(['PERIODE','DEPOT'], [st1 + '/' + st2,TobTempE.GetValue('DEPOT')], False);
    if TobTempL <> nil then
        begin
        TobTempE.PutValue('Q_CLIENT', Max(TobTempE.GetValue('Q_CLIENT'),TobTempL.GetValue('QSOR')));
        TobTempE.PutValue('Q_FOURNI', Max(TobTempE.GetValue('Q_FOURNI'),TobTempL.GetValue('QENT')));
        end;
    end;
end;

procedure CalculRupture(var TobPiece, TOBArt, TobMvts : TOB; var TobEvol : TOB);
var
    TobTempE  : Tob;
    TobTemp   : Tob;
    TobDispo  : TOB;
    Q         : TQuery;
    i_ind1    : integer;
    Select    : String;
    Annee     : Word;
    Mois      : Word;
    Jour      : Word;
begin
  //  Calcul du stock physique, tous depot ou en fonction des depot choisis
  TobDispo := TOB.Create('', nil, -1);

  Select := 'Select GQ_DEPOT,GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,GQ_PREPACLI from DISPO where GQ_ARTICLE="' +
            TOBArt.GetValue('GA_ARTICLE') + '" and GQ_CLOTURE="-" ';

  if DepDeb <> '' then Select := Select + 'and GQ_DEPOT>="' + DepDeb + '" ';

  if DepFin <> '' then Select := Select + 'and GQ_DEPOT<="' + DepFin + '" ';

  Q := OpenSQL(Select, True,-1,'',true);

  if not Q.EOF then
  begin
    TobDispo.LoadDetailDB('DISPO', '', '', Q, False);

    //EclateDate(Date, Annee, Mois, Jour);
    DecodeDate(Date, Annee, Mois, Jour);

    if VH_GC.GCPeriodeStock = 'QUI' then
    begin
      if Jour < 15 then
        Mois := Mois + 1
      else
        Mois := Mois + 2
    end
    else if VH_GC.GCPeriodeStock = 'SEM' then
      Mois := StrToInt(Format('%.2d', [NumSemaine(TobDispo.Detail[0].GetValue('GQ_DATECLOTURE'))]));

    {*
    if TousDepots then
    begin
      TobTempE := CreerFilleEvol(TobEvol, '', Annee, Mois);
      for i_ind1 := 0 to TobDispo.Detail.Count - 1 do
      begin
        TobTemp := TobDispo.Detail[i_ind1];
        TobTempE.PutValue('Q_PHYSIQUE', TobTempE.GetValue('Q_PHYSIQUE') +
                                        TobTemp.GetValue('GQ_PHYSIQUE'));
        TobTempE.PutValue('Q_CLIENT', TobTempE.GetValue('Q_CLIENT')+TobTemp.GetValue('GQ_RESERVECLI')+
                                      TobTemp.GetValue('GQ_PREPACLI'));
        TobTempE.PutValue('Q_FOURNI', TobTempE.GetValue('Q_FOURNI')+TobTemp.GetValue('GQ_RESERVEFOU'));
      end;
    end
    else
    begin
    *}
      for i_ind1 := 0 to TobDispo.Detail.Count - 1 do
      begin
        TobTemp  := TobDispo.Detail[i_ind1];
        TobTempE := CreerFilleEvol(TobEvol, TobTemp.GetValue('GQ_DEPOT'), IntToStr(Annee), IntToStr(Mois));
        TobTempE.PutValue('Q_PHYSIQUE', TobTempE.GetValue('Q_PHYSIQUE') + TobTemp.GetValue('GQ_PHYSIQUE'));
        TobTempE.PutValue('Q_CLIENT', TobTempE.GetValue('Q_CLIENT')+TobTemp.GetValue('GQ_RESERVECLI')+TobTemp.GetValue('GQ_PREPACLI'));
        TobTempE.PutValue('Q_FOURNI', TobTempE.GetValue('Q_FOURNI')+TobTemp.GetValue('GQ_RESERVEFOU'));
      end;
    //end;
  end;

  Ferme(Q);

  TobDispo.Free;

end;

procedure EnregDonnees(var TobPiece : TOB; TOBArt, TobEvol, TobMvts : TOB;TOBL : TOB;Lignestd:boolean;FromCalculDocuments : boolean);
var
    TobTemp, TobTempL, TobDepot, TobSuggest : TOB;
    Select, sSep, sDepot, sDate, sPiece, sQPhy, sCli, sFou : string;
    stRuptDepot : string;
    i_ind1 : integer;
    stkdisp, stkphy, stkcli, stkfou, stkmin, stkmax : double;
    MemoTemp : TStrings;
    Q : TQuery;
    EURO,DEV : RDevise;
    FUV,FUS,FUA : double;
    MtPa : double;
		CoefuaUs,PQQ,CoefUSUV : double;
		Tiers,Article,UA : string;

    procedure RuptureDepot;
    begin
      TobTempL := TobDepot.FindFirst(['GQ_DEPOT'], [stRuptDepot], False);
      if TobTempL <> nil then stkmin := TobTempL.GetValue('GQ_STOCKMIN') else stkmin := 0;
      if TobTempL <> nil then stkmax := TobTempL.GetValue('GQ_STOCKMAX') else stkmax := 999999999999;
      //
      if stkmax = 0 then stkmax := 999999999999;
      //
      stkdisp := stkphy - stkcli + stkfou;

      // positionnement des Quantités par detail de besoin
      if TOBL <> nil then DetermineQteAffecte (TOBL.GetValue('LIENLIGNE'),StkPhy);

      //
      if      stkdisp < 0      then stkdisp := -stkdisp + stkmin
      else if stkdisp < stkmin then stkdisp := stkmin - stkdisp
      else if (stkdisp = stkmin) and (Valeur = 'M') then
      begin
        if stkmax <> 999999999999 then stkdisp := stkmax
      end
      else
        stkdisp := -stkdisp;
    end;
    
    Procedure ChargeQuantite;
    begin

      if (Valeur = 'B') or (stkmax = 999999999999) then
      begin
        TobSuggest.PutValue('GL_QTESTOCK',    stkdisp);
        TobSuggest.PutValue('GL_QTEFACT',     stkdisp);
        TobSuggest.PutValue('GL_QTERESTE',    stkdisp);
      end
      else
      begin
        if Depart = 'P' then
        begin
          TobSuggest.PutValue('GL_QTESTOCK',    stkmax - stkphy);
          TobSuggest.PutValue('GL_QTEFACT',     stkmax - stkphy);
          TobSuggest.PutValue('GL_QTEPREVAVANC',stkmax - stkphy);
          TobSuggest.PutValue('GL_QTERESTE',    stkmax - stkphy);
        end
        else
        begin
          TobSuggest.PutValue('GL_QTESTOCK',     stkmax - (stkphy - stkcli + stkfou));
          TobSuggest.PutValue('GL_QTEFACT',      stkmax - (stkphy - stkcli + stkfou));
          TobSuggest.PutValue('GL_QTEPREVAVANC', stkmax - (stkphy - stkcli + stkfou));
          TobSuggest.PutValue('GL_QTERESTE',     stkmax - (stkphy - stkcli + stkfou));
        end;
        TobSuggest.AddChampSupValeur('STOCKMAX', stkmax, False);
        TobSuggest.AddChampSupValeur('STOCKNEC', stkdisp, False);
      end;
      //
    end;

    Procedure ChargeHistorique;
    Var i_ind2 : integer;
    begin

      if stkdisp = 0 then Exit;

      MemoTemp.Text := sDate + ';';

      MemoTemp.Add(sPiece + ';');
      MemoTemp.Add(sQPhy + ';');
      MemoTemp.Add(sCli + ';');
      MemoTemp.Add(sFou + ';');

      //  enregistrement des historiques de consommation annee n - 1
      sDepot := '';
      sDate := '';
      sCli := '';
      sSep := '';

      for i_ind2 := 0 to TobMvts.Detail.Count - 1 do
      begin
        if ((i_ind2 Div 12) Mod 2) = 0 then Continue;
        TobTempL := TobMvts.Detail[i_ind2];
        if TobTempL.GetValue('DEPOT') <> stRuptDepot then Continue;
        sDate  := sDate  + sSep + TobTempL.GetValue('PERIODE');
        sCli   := sCli   + sSep + FloatToStr(TobTempL.GetValue('QSOR'));
        sSep := ';';
      end;

      MemoTemp.Add(sDate + ';');
      MemoTemp.Add(sCli + ';');

      //  enregistrement des historiques de consommation annee n - 2
      sDepot := '';
      sDate := '';
      sCli := '';
      sSep := '';

      for i_ind2 := 0 to TobMvts.Detail.Count - 1 do
      begin
        if ((i_ind2 Div 12) Mod 2) <> 0 then Continue;
        TobTempL := TobMvts.Detail[i_ind2];
        if TobTempL.GetValue('DEPOT') <> stRuptDepot then Continue;
        sDate  := sDate  + sSep + TobTempL.GetValue('PERIODE');
        sCli   := sCli   + sSep + FloatToStr(TobTempL.GetValue('QSOR'));
        sSep := ';';
      end;

      MemoTemp.Add(sDate + ';');
      MemoTemp.Add(sCli + ';');

    end;

    procedure GenereSuggestbyLignedoc;
    begin
      //  Creation de la ligne de document associée au calcul
      Inc (NumLig);

      TobSuggest := TOB.Create('LIGNE', TobPiece, -1);

      AddLesSupLigne ( TOBSuggest,false);
      TobSuggest.InitValeurs;
      InitLesSupLigne (TOBSuggest);
      PieceVersLigne(TobPiece, TobSuggest);
      InitLigne(TOBPiece, TobArt, TobSuggest);
      //
      CoefUSUV  := TOBL.GetDouble('GL_COEFCONVQTEVTE');
      Tiers     := TOBL.GetValue('GL_FOURNISSEUR');
      Article   := TOBL.GetValue('GL_ARTICLE');

      {$IFDEF BTP}
      if FromCalculDocuments then TOBSuggest.addChampSupValeur ('LIENLIGNE',TOBL.GetValue('LIENLIGNE'),false);
      {$ENDIF}

      TobSuggest.PutValue('GL_NUMLIGNE', NumLig);
      TobSuggest.PutValue('GL_DEPOT', stRuptDepot);
      TobSuggest.PutValue('GL_BLOCNOTE', MemoTemp.Text);
      if ControleFournisseurOk ( TOBL.GetValue('GL_FOURNISSEUR')) then TOBSuggest.PutValue('GL_TIERS',TOBL.GetValue('GL_FOURNISSEUR'))
                                                                  else TOBSuggest.PutValue('GL_TIERS','');
      TOBsuggest.putValue('GL_LIBELLE',TOBL.GetValue('GL_LIBELLE'));  // On reprend le libellé de la pièce d'origine
      //
      GetInfoFromCatalog(Tiers,Article,UA,PQQ,CoefUaUs);
      //
      TOBsuggest.putValue('GL_COEFCONVQTE',CoefUaUs);
      TOBsuggest.putValue('GL_COEFCONVQTEVTE',CoefUSUV);
      TOBSuggest.PutValue('GL_QUALIFQTEACH',UA);
      //
      ChargeQuantite;
      //
      TOBSuggest.PutValue('GL_QUALIFQTESTO',TOBL.GetString('GL_QUALIFQTESTO'));
      //
      if PQQ = 0 then PQQ := 1;
      //
      FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO'));
      FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE'));
      FUA := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEACH'));
      //
      // Dans les lignes les quantitatifs et valorisations sont exprimé en UV donc passage en US
      if coefUsUv <> 0 then
      begin
        MtPa := TOBL.GetValue('GL_DPA') * CoefUSUV;
        TobSuggest.PutValue('GL_QTEPREVAVANC',TOBL.getValue('GL_QTEFACT')/CoefUSUV);
        TobSuggest.PutValue('GL_PUHTBASE',MtPa);
      end else
      begin
        MtPa := TOBL.GetValue('GL_DPA')/ FUV * FUs;
        TobSuggest.PutValue('GL_QTEPREVAVANC',TOBL.getValue('GL_QTEFACT')* FUV  / FUS );
        TobSuggest.PutValue('GL_PUHTBASE',MtPa);
      end;

      TobSuggest.PutValue('GL_PUHT', MtPA);
      TobSuggest.PutValue('GL_PUHTDEV', MtPA);

    end;

    //---------------------------------------------------------------------------------
    //
    //---------------------------------------------------------------------------------
    Procedure GenereSuggest;
    begin
      //
      TobSuggest.PutValue('GL_NUMLIGNE', NumLig);
      TobSuggest.PutValue('GL_DEPOT', stRuptDepot);
      TobSuggest.PutValue('GL_BLOCNOTE', MemoTemp.Text);
      //
      TOBSuggest.PutValue('GL_TIERS',TOBArt.GetValue('GA_FOURNPRINC')) ;
      //
      GetInfoFromCatalog(Tiers,Article,UA,PQQ,CoefUaUs);
      //
      TOBsuggest.putValue('GL_COEFCONVQTE',CoefUaUs);
      TOBsuggest.putValue('GL_COEFCONVQTEVTE',CoefUSUV);
      TOBSuggest.PutValue('GL_QUALIFQTEACH',UA);
      //
      MtPa := TobTempL.GetValue('GQ_PMAP');
      //
      TobSuggest.PutValue('GL_PUHTBASE', MtPa);
      //
      TobSuggest.PutValue('GL_PUHT',     MtPA);
      TobSuggest.PutValue('GL_PUHTDEV',  MtPA);
      //
      TobSuggest.PutValue('GL_QTEPREVAVANC',stkdisp);
      //
      ChargeQuantite;
      //
      IF LigneStd THEN Calcul_Fournisseur(TobSuggest, Prio1+';'+Prio2+';'+Prio3+';'+Prio4);
      //
    end;


begin

  FillChar(EURO,Sizeof(EURO),#0);
  EURO.Code:=V_PGI.DevisePivot ;
  GetInfosDevise(EURO) ;
  // --
  FillChar(DEV,Sizeof(DEV),#0);
  DEV.Code:=TOBPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV) ;
  //
  MemoTemp  := TStringList.Create;
  sSep      := '';
  stkphy    := 0.0;
  stkcli    := 0.0;
  stkfou    := 0.0;
  sDate     := '';
  sPiece    := '';
  sQPhy     := '';
  sCli      := '';
  sFou      := '';
  sSep      := '';

  TobDepot := TOB.Create('', nil, -1);
  if (WithStocks) and (((TOBL <> nil) AND(TOBL.GetValue('GL_TENUESTOCK')='X')) OR (TOBL=NIL)) Then
  begin
    Select:='Select * from DISPO where GQ_ARTICLE="' + TOBArt.GetValue('GA_ARTICLE') + '" and GQ_CLOTURE="-"';
    Q := OpenSQL(Select, True,-1,'',true);

    if not Q.EOF then TobDepot.LoadDetailDB('DISPO', '', '', Q, False);
    Ferme(Q);
  end;

  //TobSuggest := nil;
  TobEvol.Detail.Sort('ORDRETRI');
  for i_ind1  := 0 to TobEvol.Detail.Count - 1 do
  begin
    TobTemp   := TobEvol.Detail[i_ind1];

    if (i_ind1 = 0) then
    begin
      stkphy      := TobTemp.GetValue('Q_PHYSIQUE'); //and TousDepots
      stRuptDepot := TobTemp.GetValue('DEPOT');
    end;

    if (TobTemp.GetValue('DEPOT') <> stRuptDepot) and (stRuptDepot <> '') then
    begin
      RuptureDepot;
      //
      ChargeHistorique;

      If TOBL <> nil then
        GenereSuggestbyLignedoc    //  Creation de la ligne de document associée au calcul
      else
      begin
        Inc (NumLig);

        TobSuggest := TOB.Create('LIGNE', TobPiece, -1);

        AddLesSupLigne ( TOBSuggest,false);
        TobSuggest.InitValeurs;
        InitLesSupLigne (TOBSuggest);
        PieceVersLigne(TobPiece, TobSuggest);
        InitLigne(TOBPiece, TobArt, TobSuggest);
        //
        GenereSuggest;
      end;

      if LigneStd then Calcul_Fournisseur(TobSuggest, Prio1+';'+Prio2+';'+Prio3+';'+Prio4);
      TobSuggest.PutValue('GL_RECALCULER','X') ;
    end;
    //
    sDate  := '';
    sPiece := '';
    sQPhy  := '';
    sCli   := '';
    stkcli := 0;
    sFou   := '';
    stkfou := 0;
    sSep := '';

    sDate  := sDate  + sSep + TobTemp.GetValue('DATE');
    sPiece := sPiece + sSep + IntToStr(TobTemp.GetValue('PIECE'));
    sQPhy  := sQPhy  + sSep + StrF00(TobTemp.GetValue('Q_PHYSIQUE'), GetParamSoc('SO_DECQTE'));
    stkphy := stkphy + TobTemp.GetValue('Q_PHYSIQUE');
    sCli   := sCli   + sSep + StrF00(TobTemp.GetValue('Q_CLIENT'), GetParamSoc('SO_DECQTE'));
    stkcli := stkcli + TobTemp.GetValue('Q_CLIENT');
    sFou   := sFou   + sSep + StrF00(TobTemp.GetValue('Q_FOURNI'), GetParamSoc('SO_DECQTE'));
    stkfou := stkfou + TobTemp.GetValue('Q_FOURNI');
    sSep := ';';

    if (TobTemp.GetValue('DEPOT') <> stRuptDepot) then stkphy := TobTemp.GetValue('Q_PHYSIQUE');
    stRuptDepot := TobEvol.Detail[i_ind1].GetValue('DEPOT');
  end;

  RuptureDepot;

  ChargeHistorique;

  if stkdisp <= 0 then
  begin
    TOBDepot.Free;
    MemoTemp.Free;
    Exit;
  end;
  //
  If TOBL <> nil then
    GenereSuggestbyLignedoc
  else
  begin
    //  Creation de la ligne de document associée au calcul
    Inc (NumLig);

    TobSuggest := TOB.Create('LIGNE', TobPiece, -1);

    AddLesSupLigne ( TOBSuggest,false);
    TobSuggest.InitValeurs;
    InitLesSupLigne (TOBSuggest);
    PieceVersLigne(TobPiece, TobSuggest);
    InitLigne(TOBPiece, TobArt, TobSuggest);

    GenereSuggest;      //
  end;

  TOBDepot.Free;
  MemoTemp.Free;

end;

procedure EclateDate(DateOrig : TDateTime; var Annee, Mois, Jour : string);
var
    st1 : string;

begin
st1 := FormatDateTime ('dd' + DateSeparator + 'mm' + DateSeparator + 'yyyy', DateOrig);
Jour  := ReadTokenPipe(st1, DateSeparator);
Mois  := ReadTokenPipe(st1, DateSeparator);
Annee := ReadTokenPipe(st1, DateSeparator);
end;

function CreerFilleEvol(TobEvol : TOB; Depot, Annee, Mois : string) : TOB;
var
    TobTemp : TOB;
begin
TobTemp := TobEvol.FindFirst(['DEPOT','DATE'], [Depot, Annee + '/' + Mois], False);
if TobTemp <> nil then
    begin
    Result := TobTemp;
    Exit;
    end;
TobTemp := TOB.Create('', TobEvol, -1);
TobTemp.AddChampSup('DEPOT', True);
TobTemp.AddChampSup('DATE', True);
TobTemp.AddChampSup('PIECE', True);
TobTemp.AddChampSup('Q_PHYSIQUE', True);
TobTemp.AddChampSup('Q_CLIENT', True);
TobTemp.AddChampSup('Q_FOURNI', True);
TobTemp.AddChampSup('ORDRETRI', True);
TobTemp.PutValue('DEPOT', Depot);
TobTemp.PutValue('DATE', Annee + '/' + Mois);
TobTemp.PutValue('PIECE', 0);
TobTemp.PutValue('Q_PHYSIQUE', 0.0);
TobTemp.PutValue('Q_CLIENT', 0.0);
TobTemp.PutValue('Q_FOURNI', 0.0);
TobTemp.PutValue('ORDRETRI', Depot + Annee + Mois);
Result := TobTemp;
end;

procedure GenereEvolPeriode(TobEvol : TOB; Depot : string);
var
    st1, st2 : string;
    i_ind1, i_ind2, i_ind3, AA, MM, AADeb, MMDeb, AAFin, MMFin : integer;
    Annee, Mois, Jour : word;
    datetrav : TDateTime;

begin
//  Mise en place d'une tob evol par mois dans la periode choisie
st1 := DatDeb;
if st1 = '' then st1 := DateToStr(Date);
st2 := ReadTokenPipe(st1, '/');
MMDeb := StrToInt(ReadTokenPipe(st1, '/'));
AADeb := StrToInt(st1);
datetrav := EncodeDate(AADeb,MMDeb,1);
st1 := DatFin;
st2 := ReadTokenPipe(st1, '/');
MMFin := StrToInt(ReadTokenPipe(st1, '/'));
AAFin := StrToInt(st1);
if AADeb = AAFin then
    begin
    i_ind2 := MMFin - MMDeb;
    if VH_GC.GCPeriodeStock = 'SEM' then i_ind2 := Trunc(i_ind2 * 4 * 1.0834) // 1.0834 = 52 / 48 = nb sem par mois
    else if VH_GC.GCPeriodeStock = 'QUI' then i_ind2 := i_ind2 * 2;
    i_ind3 := 1;
    for i_ind1 := 0 to i_ind2 do
        if VH_GC.GCPeriodeStock = 'MOI' then
            CreerFilleEvol(TobEvol, Depot, IntToStr(AADeb), Format('%.2d', [MMDeb + i_ind1]))
            else
            if VH_GC.GCPeriodeStock = 'SEM' then
                begin
                CreerFilleEvol(TobEvol, Depot, IntToStr(AADeb), Format('%.2d', [NumSemaine(datetrav)]));
                datetrav := datetrav + 7;
                end
                else
                begin
                CreerFilleEvol(TobEvol, Depot, IntToStr(AADeb), Format('%.2d', [MMDeb + i_ind1]) + IntToStr(i_ind3));
                Inc(i_ind3);
                if (VH_GC.GCPeriodeStock = 'QUI') and (i_ind3 > 2) then i_ind3 := 1
                end;
    end
    else
    begin
    AA := AADeb;
    MM := MMDeb;
    i_ind2 := (12 * (AAFin - AADeb - 1)) + MMFin + (12 - MMDeb);
    if VH_GC.GCPeriodeStock = 'SEM' then i_ind2 := Trunc(i_ind2 * 4 * 1.0834) // 1.0834 = 52 / 48 = nb sem par mois
    else if VH_GC.GCPeriodeStock = 'QUI' then i_ind2 := i_ind2 * 2;
    i_ind3 := 1;
    for i_ind1 := 0 to i_ind2 do
        begin
        if VH_GC.GCPeriodeStock = 'MOI' then
            CreerFilleEvol(TobEvol, Depot, IntToStr(AA), Format('%.2d', [MM]))
            else
            if VH_GC.GCPeriodeStock = 'SEM' then
                begin
                CreerFilleEvol(TobEvol, Depot, IntToStr(AA), Format('%.2d', [NumSemaine(datetrav)]));
                datetrav := datetrav + 7;
                DecodeDate(datetrav, Annee, Mois, Jour);
                if MM <> Mois then Inc(MM);
                end
                else
                begin
                CreerFilleEvol(TobEvol, Depot, IntToStr(AA), Format('%.2d', [MM]) + IntToStr(i_ind3));
                Inc(i_ind3);
                if (VH_GC.GCPeriodeStock = 'QUI') and (i_ind3 > 2) then i_ind3 := 1
                end;
        Inc(MM);
        if MM > 12 then
            begin
            MM := 1;
            Inc(AA);
            end;
        end;
    end;
end;

function CreerFilleMvts(TobMvts : TOB; Depot, Annee, Mois : string) : TOB;
var
    TobTemp : TOB;
begin
TobTemp := TobMvts.FindFirst(['PERIODE','DEPOT'], [Annee + '/' + Mois, Depot], False);
if TobTemp = nil then
    begin
    GenereMvtsPeriode(TobMvts, Depot);
    TobTemp := TobMvts.FindFirst(['PERIODE','DEPOT'], [Annee + '/' + Mois, Depot], False);
    end;
Result := TobTemp;
end;

procedure GenereMvtsPeriode(TobMvts : TOB; Depot : string);
var
    TobTemp : TOB;
    st1, st2 : string;
    i_ind1, i_ind2, i_ind3, AA, MM, AADeb, MMDeb, AAFin, MMFin : integer;
    Annee, Mois, Jour : word;
    datetrav : TDateTime;

begin
//  Mise en place d'une tob evol par mois dans la periode choisie
st1 := DateToStr(Date);
st2 := ReadTokenPipe(st1, '/');
MMFin := StrToInt(ReadTokenPipe(st1, '/'));
AAFin := StrToInt(st1);
MMDeb := MMFin;
AADeb := AAFin - 2;
datetrav := EncodeDate(AADeb,MMDeb,1);
AA := AADeb;
MM := MMDeb;
i_ind2 := (12 * (AAFin - AADeb - 1)) + MMFin + (12 - MMDeb);
if VH_GC.GCPeriodeStock = 'SEM' then i_ind2 := Trunc(i_ind2 * 4 * 1.0834) // 1.0834 = 52 / 48 = nb sem par mois
else if VH_GC.GCPeriodeStock = 'QUI' then i_ind2 := i_ind2 * 2;
i_ind3 := 1;
for i_ind1 := 1 to i_ind2 do
    begin
    TobTemp := TOB.Create('', TobMvts, -1);
    TobTemp.AddChampSup('PERIODE', True);
    TobTemp.AddChampSup('DEPOT', True);
    TobTemp.AddChampSup('QENT', True);
    TobTemp.AddChampSup('QSOR', True);
    TobTemp.AddChampSup('COEFENT', True);
    TobTemp.AddChampSup('COEFSOR', True);
    TobTemp.AddChampSup('ORDRETRI', True);
    TobTemp.PutValue('DEPOT', Depot);
    TobTemp.PutValue('QENT', 0.0);
    TobTemp.PutValue('QSOR', 0.0);
    TobTemp.PutValue('COEFENT', 1.0);
    TobTemp.PutValue('COEFSOR', 1.0);
    if VH_GC.GCPeriodeStock = 'MOI' then
        begin
        TobTemp.PutValue('PERIODE', IntToStr(AA) + '/' + Format('%.2d', [MM]));
        TobTemp.PutValue('ORDRETRI', Depot + IntToStr(AA) + Format('%.2d', [MM]));
        Inc(MM);
        end
        else
        begin
        if VH_GC.GCPeriodeStock = 'SEM' then
            begin
            TobTemp.PutValue('PERIODE', IntToStr(AA) + '/' + Format('%.2d', [NumSemaine(datetrav)]));
            TobTemp.PutValue('ORDRETRI', Depot + IntToStr(AA) + Format('%.2d', [NumSemaine(datetrav)]));
            datetrav := datetrav + 7;
            DecodeDate(datetrav, Annee, Mois, Jour);
            if MM <> Mois then Inc(MM);
            end
            else
            begin
            TobTemp.PutValue('PERIODE', IntToStr(AA) + '/' + Format('%.2d', [MM]) + IntToStr(i_ind3));
            TobTemp.PutValue('ORDRETRI', Depot + IntToStr(AA) + Format('%.2d', [MM]) + IntToStr(i_ind3));
            Inc(i_ind3);
            if (VH_GC.GCPeriodeStock = 'QUI') and (i_ind3 > 2) then
                begin
                i_ind3 := 1;
                Inc(MM);
                end;
            end;
        end;
    if MM > 12 then
        begin
        MM := 1;
        Inc(AA);
        end;
    end;
end;

procedure Calcul_Fournisseur(TobSuggest : TOB; Param : string);
var
    TobTemp, TobTemp2, TobArt : TOB;
    Q : TQuery;
    i_ind1 : integer;
    MemoTemp : TStrings;
//    FUS, FUA : double;
    Select, Art, Crit1, Crit2, Crit3, Crit4 : string;
    Param1, Param2, Param3, Param4,UA : string;
{$IFDEF BTP}
    MtPA,CoefuaUs : double;
    fournisseur : string;
{$ENDIF}
begin
	CoefuaUs := 0;
  MemoTemp := TStringList.Create;
  Param1 := ReadTokenSt(Param);
  Param2 := ReadTokenSt(Param);
  Param3 := ReadTokenSt(Param);
  Param4 := ReadTokenSt(Param);
  //  Calcul du fournisseur eventuel en fonction des priorités.
  Art := TobSuggest.GetValue('GL_ARTICLE');
  MemoTemp.Text := TobSuggest.GetValue('GL_BLOCNOTE');
  for i_ind1 := MemoTemp.Count - 1 downto 0 do
  begin
    Select := MemoTemp.Strings[i_ind1];
    if Pos('Four', Select) <> 0 then MemoTemp.Delete(i_ind1);
  end;
  TobArt := TOB.Create('', nil, -1);
  Select := 'Select GA_ARTICLE,GA_FOURNPRINC,GA_DPA,GA_DPR,GA_PMRP,GA_PMAP,GA_PAHT,GA_PRHT,GA_COEFFG,GA_DPRAUTO,GA_CALCPRIXPR,' +
            'GA_PVHT,GA_PVTTC,GA_VALLIBRE1,GA_VALLIBRE2,GA_VALLIBRE3,GA_CALCAUTOHT,GA_COEFCALCHT,GA_CALCPRIXHT from ARTICLE '+
            'WHERE GA_ARTICLE="' + Art + '"';
  Q := OpenSQL(Select, True,-1,'',true);
  if Q.EOF then
  begin
    ferme (Q);
    TobArt.Free;
    //    TobArt := nil;
    MemoTemp.Free;
    Exit;
  end;

  TobArt.SelectDB('', Q);
  Ferme(Q);

  Select := 'Select * from DISPO where GQ_ARTICLE="' + Art + '" and GQ_CLOTURE="-"';
  TobArt.LoadDetailDBFromSQL('DISPO', Select);

  // on charge le catalogue eventuel trié en fct des criteres de priorite
  if Param1 = 'APP' then Crit1 := 'GCA_DELAILIVRAISON';
  if Param1 = 'FPR' then Crit1 := 'GCA_TIERS';
  if Param1 = 'COT' then Crit1 := 'GCA_COTE';
  if Param1 = 'MPT' then Crit1 := 'GCA_DPA';
  if Param2 = 'APP' then Crit2 := 'GCA_DELAILIVRAISON';
  if Param2 = 'FPR' then Crit2 := 'GCA_TIERS';
  if Param2 = 'COT' then Crit2 := 'GCA_COTE';
  if Param2 = 'MPT' then Crit2 := 'GCA_DPA';
  if Param3 = 'APP' then Crit3 := 'GCA_DELAILIVRAISON';
  if Param3 = 'FPR' then Crit3 := 'GCA_TIERS';
  if Param3 = 'COT' then Crit3 := 'GCA_COTE';
  if Param3 = 'MPT' then Crit3 := 'GCA_DPA';
  if Param4 = 'APP' then Crit4 := 'GCA_DELAILIVRAISON';
  if Param4 = 'FPR' then Crit4 := 'GCA_TIERS';
  if Param4 = 'COT' then Crit4 := 'GCA_COTE';
  if Param4 = 'MPT' then Crit4 := 'GCA_DPA';

  TobTemp := TOB.Create('', nil, -1);
  Select := 'Select GCA_ARTICLE,GCA_TIERS,GCA_DPA,GCA_QUALIFUNITEACH,GCA_PRIXBASE,GCA_DELAILIVRAISON,' +
            'GCA_COTE,GCA_REFERENCE,GCA_PRIXPOURQTEAC,GCA_COEFCONVQTEACH from CATALOGU ' +
            'where GCA_ARTICLE="' + Art + '" and GCA_DATESUP>="' + USDateTime(Date) +
            '" order by ' + Crit1;
  if Crit2 <> '' then Select := Select + ',' + Crit2;
  if Crit3 <> '' then Select := Select + ',' + Crit3;
  if Crit4 <> '' then Select := Select + ',' + Crit4;
  Q := OpenSQL(Select, True,-1,'',true);
  if not Q.EOF then TobTemp.LoadDetailDB('CATALOGU', '', '', Q, False);
  Ferme(Q);
  // Rien dans le catalogue, y-a-t'il un fourn. principal ?
  if TobTemp.Detail.Count = 0 then
  begin
    if (Pos('FPR', Param1+Param2+Param3+Param4) <> 0) and (TobArt.GetValue('GA_FOURNPRINC') <> '') then
    begin
      TobSuggest.PutValue('GL_TIERS', TobArt.GetValue('GA_FOURNPRINC'));
      TobSuggest.PutValue('GL_PUHT', QuelPrixBase(TobSuggest.GetValue('GL_NATUREPIECEG'),
            TobSuggest.GetValue('GL_DEPOT'),TobArt,TobSuggest));
      //        TobSuggest.PutValue('GL_PUHT', TobArt.GetValue('GA_DPA'));
    end;
    TobArt.Free;
    //    TobArt := nil;
    TobTemp.Free;
    //    TobTemp := nil;
    MemoTemp.Free;
    Exit;
  end;
  // Le catalog existe. on le tri en fonction des priorités définies.
  // si le premier critere de tri est le fournisseur principal et que celui ci existe
  // il annule implicitement tous les autres et on le choisit.
  // sinon on prend le premier enregistrement lu dans le catalogue
  if (Param1 = 'FPR') and (TobArt.GetValue('GA_FOURNPRINC') <> '') then
  begin
    TobTemp2 := TobTemp.FindFirst(['GCA_TIERS', 'GCA_ARTICLE'],
    [TobArt.GetValue('GA_FOURNPRINC'), TobArt.GetValue('GA_ARTICLE')], True);
    if TobTemp2 <> nil then
    begin
      TobArt.AddChampSupValeur('GCA_PRIXBASE', TobTemp2.GetValue('GCA_PRIXBASE'));
      TobArt.AddChampSupValeur('GCA_DPA', TobTemp2.GetValue('GCA_DPA'));
      TobSuggest.PutValue('GL_QUALIFQTEACH',TOBTemp2.getValue('GCA_QUALIFUNITEACH'));
      TobSuggest.PutValue('GL_COEFCONVQTE',TOBTemp2.getValue('GCA_COEFCONVQTEACH'));
    end;
    TobSuggest.PutValue('GL_TIERS', TobArt.GetValue('GA_FOURNPRINC'));
{$IFDEF BTP}
    Fournisseur := TobArt.GetValue('GA_FOURNPRINC');
    MtPa := RecupTarifAch (Fournisseur,Art,Ua,CoefUaUs,Turstock);
    TobSuggest.PutValue('GL_PUHT', MtPa);
    TobSuggest.PutValue('GL_PUHTDEV', MtPa);
{$ELSE}
    TobSuggest.PutValue('GL_PUHT', QuelPrixBase(TobSuggest.GetValue('GL_NATUREPIECEG'),
    TobSuggest.GetValue('GL_DEPOT'),TobArt,TobSuggest));
{$ENDIF}
    //  TobSuggest.PutValue('GL_PUHT', TobArt.GetValue('GA_DPA'));
  end else
  begin
  // on va chercher le meilleur prix tarif de chaque fournisseur pour l'article
    for i_ind1 := 0 to TobTemp.Detail.Count - 1 do
    begin
      TobTemp2 := TobTemp.Detail[i_ind1];
      TobTemp2.AddChampSup('ORDRETRI', True);
      if Pos('MPT', Param1+Param2+Param3+Param4) <> 0 then
      begin
        Recherche_Tarif(TobArt, TobSuggest, TobTemp2)
      end
      else
      begin
      {$IFDEF BTP}
        Fournisseur := TobTemp2.GetValue('GCA_TIERS');
        MtPa := RecupTarifAch (Fournisseur,Art,Ua,CoefuaUs,TurStock );
        TobArt.AddChampSupValeur('GCA_PRIXBASE', mtPa);
        TobArt.AddChampSupValeur('GCA_DPA', MtPa);
        TobTemp2.PutValue('GCA_DPA', mtPa);
        {$ELSE}
        TobArt.AddChampSupValeur('GCA_PRIXBASE', TobTemp2.GetValue('GCA_PRIXBASE'));
        TobArt.AddChampSupValeur('GCA_DPA', TobTemp2.GetValue('GCA_DPA'));
        TobTemp2.PutValue('GCA_DPA', QuelPrixBase(TobSuggest.GetValue('GL_NATUREPIECEG'),
              TobSuggest.GetValue('GL_DEPOT'),TobArt,TobSuggest));
        {$ENDIF}
      end;
      Select := TobTemp2.GetValue(Crit1);
      if Crit2 <> '' then Select := Select + ';' + FloatToStr(TobTemp2.GetValue(Crit2)) else Select := Select + ';';
      if Crit3 <> '' then Select := Select + ';' + TobTemp2.GetValue(Crit3) else Select := Select + ';';
      if Crit4 <> '' then Select := Select + ';' + TobTemp2.GetValue(Crit4) else Select := Select + ';';
      TobTemp2.PutValue('ORDRETRI',Select);
    end;
    if Pos('MPT', Param1+Param2+Param3+Param4) <> 0 then TobTemp.Detail.Sort('GCA_DPA');
    TobSuggest.PutValue('GL_TIERS', TobTemp.Detail[0].GetValue('GCA_TIERS'));

    TobSuggest.PutValue('GL_PUHT', TobTemp.Detail[0].GetValue('GCA_DPA'));
    TobSuggest.PutValue('GL_PUHTDEV', TobTemp.Detail[0].GetValue('GCA_DPA'));

    TobSuggest.PutValue('GL_REFARTTIERS', TobTemp.Detail[0].GetValue('GCA_REFERENCE'));
    TobSuggest.PutValue('GL_QUALIFQTEACH',TobTemp.Detail[0].GetValue('GCA_QUALIFUNITEACH')) ;
    (*
    FUS := RatioMesure('PIE', TobSuggest.GetValue('GL_QUALIFQTESTO'));
    FUA := RatioMesure('PIE', TobSuggest.GetValue('GL_QUALIFQTEACH'));
    *)
    (* Va falloir l'expliquer caaaaaaaa
    TobSuggest.PutValue('GL_QTEFACT', (TobSuggest.GetValue('GL_QTESTOCK') * FUS) / FUA);
    TobSuggest.PutValue('GL_QTESTOCK',TobSuggest.GetValue('GL_QTEFACT'));
    TobSuggest.PutValue('GL_QTERESTE',TobSuggest.GetValue('GL_QTEFACT'));
    *)
    MemoTemp.Add(' ');
    for i_ind1 := 0 to TobTemp.Detail.Count - 1 do
    begin
      MemoTemp.Add('Four : ' +
      TobTemp.Detail[i_ind1].GetValue('GCA_TIERS') + ';' +
      TobTemp.Detail[i_ind1].GetValue('GCA_REFERENCE') + ';' +
      FloatToStr(TobTemp.Detail[i_ind1].GetValue('GCA_PRIXBASE')) + ';' +
      TobTemp.Detail[i_ind1].GetValue('GCA_COTE') + ';' +
      IntToStr(TobTemp.Detail[i_ind1].GetValue('GCA_DELAILIVRAISON')));
    end;
    TobSuggest.PutValue('GL_BLOCNOTE', MemoTemp.Text);
  end;
  TobArt.Free;
  //TobArt := nil;
  TobTemp.Free;
  //TobTemp := nil;
  MemoTemp.Free;
end;

procedure Recherche_Tarif(TobArt, TobSuggest, TobCat : TOB);
var
    TobTiers, TobPiece, TobLigne, TobTarif : TOB;
    TobLigneTarif: TOB;
    DEV : RDEVISE;
{$IFDEF BTP}
    MTPAF,CoefUaUs : double;
    UA : string ;
//    FUA,FUS,PQQ : double;
{$ENDIF}
begin
CoefUaUs := 0;
FillChar(DEV,Sizeof(DEV),#0);
DEV.Code:=V_PGI.DevisePivot ;
GetInfosDevise(DEV) ;
TobPiece := TOB.Create('PIECE', nil, -1);
TobLigne := TOB.Create('LIGNE', nil, -1);
TobLigneTarif  := TOB.Create('LIGNETARIF' , nil, -1);
TobTarif := TOB.Create('TARIF', nil, -1);
TobTiers := TOB.Create('TIERS', nil, -1);
TobPiece.InitValeurs;
TobLigne.InitValeurs;
TobTarif.InitValeurs;
TobTiers.InitValeurs;
TobTiers.PutValue('T_TIERS', TobCat.GetValue('GCA_TIERS'));
TobTiers.LoadDB;
TobPiece.PutValue('GP_NATUREPIECEG', 'CF');
TobPiece.PutValue('GP_TIERS', TobCat.GetValue('GCA_TIERS'));
TOBPiece.PutValue('GP_DATEPIECE',V_PGI.DateEntree);
TOBPiece.PutValue('GP_DATECREATION',V_PGI.DateEntree);
TOBPiece.PutValue('GP_DATEMODIF',V_PGI.DateEntree);
TOBPiece.PutValue('GP_DEVISE',V_PGI.DevisePivot);
TOBPiece.PutValue('GP_FACTUREHT', 'X');
TOBPiece.PutValue('GP_VENTEACHAT',GetInfoParPiece('CF','GPP_VENTEACHAT')) ;
TOBPiece.PutValue('GP_DEPOT', '');
PieceVersLigne(TobPiece, TobLigne);
AddLesSupLigne(TobLigne,False) ;
TobLigne.PutValue('GL_ARTICLE', TobArt.GetValue('GA_ARTICLE'));
TobLigne.PutValue('GL_DEPOT', TobSuggest.GetValue('GL_DEPOT'));
TobLigne.PutValue('GL_QTEFACT', TobSuggest.GetValue('GL_QTESTOCK'));
TobLigne.PutValue('GL_QTESTOCK', TobLigne.GetValue('GL_QTEFACT'));
TobLigne.PutValue('GL_QTERESTE', TobLigne.GetValue('GL_QTEFACT'));
{$IFNDEF BTP}
TobLigne.PutValue('GL_PUHTDEV', TobSuggest.GetValue('GL_PUHT'));
TarifVersLigne ( TobArt, TobTiers, TobLigne, TobLigneTarif, TobPiece, TobTarif, True, True, DEV);
TobCat.PutValue('GCA_DPA', TobLigne.GetValue('GL_PUHT'));
{$ELSE}
MTPAF := RecupTarifAch (string(TobCat.GetValue('GCA_TIERS')) ,
                        string(TOBART.GetValue('GA_ARTICLE')),Ua,CoefuaUs,
                        TurStock);
TobLigne.PutValue('GL_PUHT',MTPAF); TobLigne.PutValue('GL_PUHTDEV',MTPAF);
TobCat.PutValue('GCA_DPA', MTPAF);
{$ENDIF}
TobPiece.Free;
TobLigne.Free;
TobTarif.Free;
TobLigneTarif.Free;
TobTiers.Free;
//TobPiece := nil;
//TobLigne := nil;
//TobTarif := nil;
//TobLigneTarif := nil;
//TobTiers := nil;
end;

Procedure InitLigne ( TOBPiece,TOBA,TOBL : TOB ) ;
Var NaturePiece,Depot : String ;
    Prix     : Double ;
    i,j               : integer ;
    RefUnique,VenteAchat : String ;
BEGIN
if TOBA=Nil then Exit ;
RefUnique:=TOBA.GetValue('GA_ARTICLE') ;
if RefUnique<>'' then
   BEGIN
   VenteAchat:=GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT') ;
   TOBL.PutValue('GL_ARTICLE',TOBA.GetValue('GA_ARTICLE')) ;
   TOBL.PutValue('GL_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE')) ;
   TOBL.PutValue('GL_REFARTSAISIE',TOBA.GetValue('GA_CODEARTICLE')) ;
   TOBL.PutValue('GL_LIBELLE',TOBA.GetValue('GA_LIBELLE')) ;
   TOBL.PutValue('GL_PRIXPOURQTE',1) ;
   TOBL.PutValue('GL_LIBCOMPL',TOBA.GetValue('GA_LIBCOMPL')) ;
   TOBL.PutValue('GL_REFARTBARRE',TOBA.GetValue('GA_CODEBARRE')) ;
   TOBL.PutValue('GL_ESCOMPTABLE',TOBA.GetValue('GA_ESCOMPTABLE')) ;
   TOBL.PutValue('GL_REMISABLEPIED',TOBA.GetValue('GA_REMISEPIED')) ;
   TOBL.PutValue('GL_REMISABLELIGNE',TOBA.GetValue('GA_REMISELIGNE')) ;
   TOBL.PutValue('GL_TENUESTOCK',TOBA.GetValue('GA_TENUESTOCK')) ;
   TOBL.PutValue('GL_TARIFARTICLE',TOBA.GetValue('GA_TARIFARTICLE')) ;
   TOBL.PutValue('GL_TYPEREF','ART') ;
   TOBL.PutValue('GL_TYPELIGNE','ART') ;
   {Familles, collection, domaine}
   TOBL.PutValue('GL_FAMILLENIV1',TOBA.GetValue('GA_FAMILLENIV1')) ;
   TOBL.PutValue('GL_FAMILLENIV2',TOBA.GetValue('GA_FAMILLENIV2')) ;
   TOBL.PutValue('GL_FAMILLENIV3',TOBA.GetValue('GA_FAMILLENIV3')) ;
   TOBL.PutValue('GL_COLLECTION',TOBA.GetValue('GA_COLLECTION')) ;
   TOBL.PutValue('GL_DOMAINE',TOBA.GetValue('GA_DOMAINE')) ;
   {Nomenclature}
   TOBL.PutValue('GL_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE')) ;
   TOBL.PutValue('GL_TYPENOMENC',TOBA.GetValue('GA_TYPENOMENC')) ;
   NaturePiece:=TOBL.GetValue('GL_NATUREPIECEG') ; Depot:=TOBL.GetValue('GL_DEPOT') ;
   for i:=1 to 5 do
       TOBL.PutValue('GL_FAMILLETAXE'+IntToStr(i),TOBA.GetValue('GA_FAMILLETAXE'+IntToStr(i))) ;
   {Unités de mesure}
   TOBL.PutValue('GL_QUALIFSURFACE',TOBA.GetValue('GA_QUALIFSURFACE')) ;
   TOBL.PutValue('GL_QUALIFVOLUME',TOBA.GetValue('GA_QUALIFVOLUME')) ;
   TOBL.PutValue('GL_QUALIFPOIDS',TOBA.GetValue('GA_QUALIFPOIDS')) ;
   TOBL.PutValue('GL_QUALIFLINEAIRE',TOBA.GetValue('GA_QUALIFLINEAIRE')) ;
   TOBL.PutValue('GL_QUALIFHEURE',TOBA.GetValue('GA_QUALIFHEURE')) ;
   TOBL.PutValue('GL_SURFACE',TOBA.GetValue('GA_SURFACE')) ;
   TOBL.PutValue('GL_VOLUME',TOBA.GetValue('GA_VOLUME')) ;
   TOBL.PutValue('GL_POIDSBRUT',TOBA.GetValue('GA_POIDSBRUT')) ;
   TOBL.PutValue('GL_POIDSNET',TOBA.GetValue('GA_POIDSNET')) ;
   TOBL.PutValue('GL_POIDSDOUA',TOBA.GetValue('GA_POIDSDOUA')) ;
   TOBL.PutValue('GL_LINEAIRE',TOBA.GetValue('GA_LINEAIRE')) ;
//   TOBL.PutValue('GL_HEURE',TOBA.GetValue('GA_HEURE')) ;
   TOBL.PutValue('GL_QUALIFQTESTO',TOBA.GetValue('GA_QUALIFUNITESTO')) ;
   TOBL.PutValue('GL_QUALIFQTEVTE',TOBA.GetValue('GA_QUALIFUNITEVTE')) ;
   {Tables libres}
   for j:=1 to 9 do
       TOBL.PutValue('GL_LIBREART'+IntToStr(j),TOBA.GetValue('GA_LIBREART'+IntToStr(j))) ;
   TOBL.PutValue('GL_LIBREARTA',TOBA.GetValue('GA_LIBREARTA')) ;
   {Affaires}
   if ctxAffaire in V_PGI.PGIContexte then
       TOBL.PutValue('GL_BLOCNOTE',TOBA.GetValue('GA_COMMENTAIRE')) ;
   {Prix}
   Prix:=QuelPrixBase(NaturePiece,Depot,TOBA,TOBL) ;
   if TOBL.GetValue('GL_FACTUREHT')='X' then
      BEGIN
      TOBL.PutValue('GL_PUHT',Prix) ; TOBL.PutValue('GL_PUHTDEV',Prix) ;
      END else
      BEGIN
      TOBL.PutValue('GL_PUTTC',Prix) ; TOBL.PutValue('GL_PUTTCDEV',Prix) ;
      END ;
   AffectePrixValo(TOBL,TOBA) ;
   {Bloc-Notes}
   if GetInfoParPiece(NaturePiece,'GPP_BLOBART')='X' then
       TOBL.PutValue('GL_BLOCNOTE',TOBA.GetValue('GA_BLOCNOTE')) ;
   {Divers}
   TOBL.PutValue('GL_DATELIVRAISON',TOBPiece.GetValue('GP_DATELIVRAISON')) ;
   TOBL.PutValue('GL_CAISSE',TOBPiece.GetValue('GP_CAISSE')) ;
   {Mode}
   if (ctxMode in V_PGI.PGIContexte) and (VenteAchat='VEN') then
       TOBL.PutValue('GL_FOURNISSEUR',TOBA.GetValue('GA_FOURNPRINC')) ;
   END ;
END ;

procedure CalculeThisDocument(TOBPiece : TOB);
var TOBBases,TOBBasesL : TOB;
    DEV : RDevise;
begin
  FillChar(DEV,Sizeof(DEV),#0);
  DEV.Code:=V_PGI.DevisePivot ;
  TobBases := TOB.Create('BASES',Nil,-1);
  TobBasesL := TOB.Create('LES BASES',Nil,-1);
  TRY
    GetInfosDevise(DEV);
    CalculFacture (nil,TOBPiece,nil,nil,nil,nil,TobBases,TOBBasesL,nil,nil,nil,nil,nil,nil,DEV);
  FINALLY
    TobBases.Free; TobBasesL.Free;
  END;
end;

function TraitDocuments (TOBpiece : TOB) : boolean;
var TOBTemp : TOB;
    indice,NumLigne : integer;
    NbrAtraiter : integer;
begin
  result := false;
  NumLigne := 0;
  NbrATraiter := TOBDetailException.detail.count + TOBDetailStd.detail.count;
  if NbrAtraiter > 0 then
  begin
    InitMove(NbrAtraiter, '');
    NumLig := 0;
    for Indice := 0 to TobDetailStd.Detail.Count - 1 do
    begin
      TobTemp := TobDetailStd.Detail[Indice];
      if TraiteUneLigneBesoin(TOBPiece, TobTemp,true,NumLigne) then result := true;
      MoveCur(False);
    end;
    for Indice := 0 to TOBDetailException.Detail.Count - 1 do
    begin
      TobTemp := TOBDetailException.Detail[Indice];
      if TraiteUneLigneBesoin(TOBPiece, TobTemp,false,NumLigne) then result := true;
      MoveCur(False);
    end;
    FiniMove();
    NumeroteLignesGC (nil,TOBPiece);
    CalculeThisDocument(TOBPiece);
{$IFDEF BTP}
    MiseEnPlaceLienReappro (TOBPiece,TOBLien);
    PurgeLesLiens (TOBLien);
{$ENDIF}
    FResulSuggestion.GZZ_SUGGEST.Text := TOBPiece.GetValue('GP_NUMERO');
  end;
end;


function RechercheDocuments (TOBLIsteDocuments : TOB) : boolean;
var Select: string;
    i_ind1 : integer;
    QQ : TQuery;
    TOBTmp: TOB;
begin
  result := false;
  i_ind1 := 0;
  if TOBListeDocuments.Detail.Count > 1 then
  begin
    Select := 'And GL_NUMERO in (';
    for i_ind1 := 0 to TOBListeDocuments.Detail.Count - 1 do
    begin
      if i_ind1 = 0 then
        Select := Select + IntToStr(TOBListeDocuments.Detail[i_ind1].GetValue('GP_NUMERO'))
      else
        Select := Select + ', ' + IntToStr(TOBListeDocuments.Detail[i_ind1].GetValue('GP_NUMERO'));
    end;
    Select := Select + ') ';
  end else if TOBListeDocuments.Detail.Count = 1 then
    Select := ' And GL_NUMERO=' + IntToStr(TOBListeDocuments.Detail[i_ind1].GetValue('GP_NUMERO')) + ' '
  else
    Select := '';
  TOBTmp := TOB.create ('LES LIGNES',nil,-1);
  TRY
    select := ConstitueRequeteDoc (Select,true);
    QQ := OpenSql (Select,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBTmp.LoadDetailDB ('LIGNE','','',QQ,false,true);
      result := true;
    end;
    ferme (QQ);
    FaisLeTriStd (TOBTmp);
  FINALLY
    TOBTMP.free;
  END;
end;

function TraiteUneLigneBesoin (TobPiece,TOBL : TOB; LigneStandard : boolean ; var Ligne : integer) : boolean;
var TOBArt,TOBMvts,TOBEvol : TOB;
    Article : string;
begin
//  Result := false;
  TOBArt := TOB.Create ('ARTICLE',nil,-1);
  TobMvts := TOB.Create('', nil, -1);
  TobEvol := TOB.Create('', nil, -1);
  TRY
    Article := copy (TOBL.GetValue('GL_ARTICLE'),1,18);
    TOBArt.putValue('GA_ARTICLE',TOBL.GetValue('GL_ARTICLE'));
    TOBART.LoadDB;
    // Pour prendre en compte la particularite de la commande
    TOBART.putValue ('GA_TENUESTOCK',TOBL.GetValue('GL_TENUESTOCK'));
    // --
    CalculHistoStock(TOBArt, TobMvts,LigneStandard);
    CalculStockPhysique(TOBArt, TobEvol);
    CalculEvolutionStock(TOBArt, TobMvts, TobEvol,LigneStandard,TOBL,True,TOBL.GetValue('GL_LIBELLE'));
    if (TobEvol.Detail.Count > 0) or (TobMvts.Detail.Count > 0) then
    begin
        EnregDonnees(TobPiece, TOBArt, TobEvol, TobMvts,TOBL,LigneStandard,True)
    end;
  FINALLY
    TOBArt.free;
    TobMvts.free;
    TobEvol.free;
    Result := true;
  END;
end;

procedure DefiniListeDocuments (TOBListeDocuments : TOB ; All : boolean);
var Select : string;
    QQ : Tquery;
    TOBTmp,TOBNew : TOB;
    Indice : integer;
begin
  Select := 'SELECT DISTINCT GP_NATUREPIECEG,GP_VIVANTE,GP_SOUCHE,GP_NUMERO,GP_INDICEG,T_LIBELLE FROM PIECE ' +
            'left join LIGNE on (GL_NATUREPIECEG=GP_NATUREPIECEG) and (GL_NUMERO=GP_NUMERO) and (GL_INDICEG=GP_INDICEG) ' +
            'left join TIERS on (T_TIERS=GP_TIERS)  ' +
            'WHERE GP_NATUREPIECEG IN ("'+ GetNaturePieceCde+'") ' ;
  if (ArtDeb <> '') and (ArtFin <> '') then
      Select := Select + 'AND (GL_CODEARTICLE>="' + ArtDeb + '" AND GL_CODEARTICLE<="' + ArtFin + '") '
  else if ArtDeb <> '' then
      Select := Select + 'AND GL_CODEARTICLE>="' + ArtDeb + '" '
  else if ArtFin <> '' then
      Select := Select + 'AND GL_CODEARTICLE<="' + ArtFin + '" ';
  if Fam1 <> '' then
      if Fam1_ <> '' then
          Select := Select + 'AND (GL_FAMILLENIV1>="' + Fam1 + '" AND GL_FAMILLENIV1<="' + Fam1_ + '") '
          else
          Select := Select + 'AND GL_FAMILLENIV1>="' + Fam1 + '" ';
  if Fam2 <> '' then
      if Fam2_ <> '' then
          Select := Select + 'AND (GL_FAMILLENIV2>="' + Fam2 + '" AND GL_FAMILLENIV2<="' + Fam2_ + '") '
          else
          Select := Select + 'AND GL_FAMILLENIV2="' + Fam2 + '" ';
  if Fam3 <> '' then
      if Fam3_ <> '' then
          Select := Select + 'AND (GL_FAMILLENIV3>="' + Fam3 + '" AND GL_FAMILLENIV3<="' + Fam3_ + '") '
          else
          Select := Select + 'AND GL_FAMILLENIV3="' + Fam3 + '" ';
  if DepDeb  <> '' then
      if DepFin  <> '' then
          Select := Select + 'AND (GL_DEPOT>="' + DepDeb + '" AND GL_DEPOT<="' + DepFin + '") '
          else
          Select := Select + 'AND GL_DEPOT="' + DepDeb + '" ';
  Select := Select + 'AND GL_DATELIVRAISON>="' + USDateTime(strtodate(DatDeb)) + '" ';
  Select := Select + 'AND GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' ;
  Select := Select + 'AND GL_TYPEDIM<>"GEN" ' ;
  Select := Select + 'AND GL_TYPELIGNE="ART" ' ;
  Select := Select + 'AND GP_VIVANTE="X" ' ;
  Select := Select + 'AND GL_VIVANTE="X" ' ;
  QQ := OpenSql (Select,true,-1,'',true);
  if not QQ.eof then
  begin
    if all then // chargement de toutes les zones de la piece
    begin
      TOBTmp := TOB.Create ('THE TEMPORAIRE',nil,-1);
      TRY
        TOBTmp.loaddetailDB ('PIECE','','',QQ,false,true);
        for Indice := 0 to TOBTmp.detail.count -1 do
        begin
          TOBNew := TOB.create ('PIECE',TOBListeDocuments,-1);
          TOBnew.dupliquer (TOBTmp.detail[Indice],false,true);
          TOBNew.loaddb(true);
        end;
      FINALLY
        TOBTmp.free;
      END;
    end else
    begin
      TOBListeDocuments.LoadDetailDB ('PIECE','','',QQ,false,true);
    end;
  end;
  ferme (QQ);
end;

procedure AjouteLigneLien (TOBL,TOBDest : TOB; Indice : integer);
var TOBLI : TOB;
begin
  TOBLI := TOB.Create ('LIENDEVCHA',TOBLien,-1);
  TOBLI.addChampSupValeur ('USED','-',false);
  TOBLI.addChampSupValeur ('LIENLIGNE',Indice,false);
  TOBLI.addChampSupValeur ('ARTICLE',TOBL.GetValue('GL_ARTICLE'),false);
  TOBLI.addChampSupValeur ('LIBELLE',TOBL.GetValue('GL_LIBELLE'),false);
  TOBLI.addChampSupValeur ('QTE',TOBL.GetValue('GL_QTEFACT'),false);
{$IFDEF BTP}
  TOBLI.PutValue('BDA_REFD',EncodeLienDevCHA(TOBL));
{$ENDIF}
  TOBLI.PutValue('BDA_NUMLD',TOBL.getValue('GL_NUMLIGNE'));
  TOBDest.addChampSupValeur ('LIENLIGNE',Indice,false);
end;

function ConstitueRequeteDoc (sWhere : string; Detail : boolean=false) : string;
var ChaineSelect : string;
begin
  ChaineSelect := ' * ';
  if not Detail then ChaineSelect :=' DISTINCT GL_ARTICLE ';
  Result := 'SELECT'+ ChaineSelect+ ', GP_VIVANTE FROM LIGNE LEFT OUTER JOIN PIECE ON (GP_NATUREPIECEG=GL_NATUREPIECEG) AND '+
            '(GP_SOUCHE=GL_SOUCHE) AND (GP_NUMERO=GL_NUMERO) AND (GP_INDICEG=GL_INDICEG) ';
  result := Result + 'WHERE GL_NATUREPIECEG IN ("'+GetNaturePieceCde+'") ' ;
  if (ArtDeb <> '') and (ArtFin <> '') then
      Result := Result + 'AND (GL_CODEARTICLE>="' + ArtDeb + '" AND GL_CODEARTICLE<="' + ArtFin + '") '
  else if ArtDeb <> '' then
      Result := Result + 'AND GL_CODEARTICLE>="' + ArtDeb + '" '
  else if ArtFin <> '' then
      Result := Result + 'AND GL_CODEARTICLE<="' + ArtFin + '" ';
  if Fam1 <> '' then
      if Fam1_ <> '' then
          Result := Result + 'AND (GL_FAMILLENIV1>="' + Fam1 + '" AND GL_FAMILLENIV1<="' + Fam1_ + '") '
          else
          Result := Result + 'AND GL_FAMILLENIV1>="' + Fam1 + '" ';
  if Fam2 <> '' then
      if Fam2_ <> '' then
          Result := Result + 'AND (GL_FAMILLENIV2>="' + Fam2 + '" AND GL_FAMILLENIV2<="' + Fam2_ + '") '
          else
          Result := Result + 'AND GL_FAMILLENIV2="' + Fam2 + '" ';
  if Fam3 <> '' then
      if Fam3_ <> '' then
          Result := Result + 'AND (GL_FAMILLENIV3>="' + Fam3 + '" AND GL_FAMILLENIV3<="' + Fam3_ + '") '
          else
          Result := Result + 'AND GL_FAMILLENIV3="' + Fam3 + '" ';
  if DepDeb  <> '' then
      if DepFin  <> '' then
          Result := Result + 'AND (GL_DEPOT>="' + DepDeb + '" AND GL_DEPOT<="' + DepFin + '") '
          else
          Result := Result + 'AND GL_DEPOT="' + DepDeb + '" ';
  Result := Result + 'AND GL_DATELIVRAISON>="' + USDateTime(strtodate(DatDeb)) + '" ';
  Result := Result + 'AND GL_DATELIVRAISON<="' + USDateTime(StrToDate(DatFin)) + '" ' ;
  Result := Result + 'AND GL_TYPEDIM<>"GEN" ' ;
  Result := Result + 'AND GL_TYPELIGNE="ART" ' ;
//  Result := Result + 'AND GL_QTERESTE > 0 ' ;
  // --- GUINIER ---
  Result := Result + 'AND (GL_QTERESTE <> 0) ' ;
{$IFNDEF BTP}
  Result := Result + 'AND GL_TYPEARTICLE<>"PRE" ' ; // On enleve les prestations
{$ENDIF}
  Result := Result + 'AND GL_VIVANTE="X" ' ;
  Result := Result + 'AND GP_VIVANTE="X" ' ;
  if SWhere <> '' then Result := Result+Swhere;  // prise en compte de la liste des documents a traiter
  Result := Result + 'ORDER BY GL_ARTICLE,GL_DEPOT';
end;

procedure AjouteCumule (TOBDest: TOB; var NumLigRea : integer; TOBL : TOb;ModifOk:boolean=true);
var  TOBNL,TOBF{,TOBLi}: TOb;
     Requete,RefOrigine : string;
     QteRecep : double;
     QQ : Tquery;
     FUA,FUS,FUV,CoefUaUs : double;
     Qte : double;
//     Indice : integer;
begin
  // Modif Suite Gestion Reliquat
  FUV := RatioMesure('PIE', TOBL.GetValue('GL_QUALIFQTEVTE'));
  FUS := RatioMesure('PIE', TOBL.GEtVAlue('GL_QUALIFQTESTO'));

  QteRecep := 0;
  RefOrigine := EncodeRefPiece(TOBL);
  // Requete en group by bicose le GL_QUALIFQTEACH depend du fournisseur
  requete := 'SELECT SUM(GL_QTERESTE) AS QTERECEP, SUM(GL_MTRESTE) AS MTRECEP, GL_QUALIFQTEACH,GL_COEFCONVQTE FROM LIGNE WHERE GL_PIECEORIGINE="'+RefOrigine+
             '" AND GL_NATUREPIECEG IN ('+GetPieceAchat(True,True)+') AND GL_VIVANTE="X" GROUP BY GL_QUALIFQTEACH,GL_COEFCONVQTE';
  QQ := OpenSql (requete,true,-1,'',true);
  TRY
    if not QQ.eof then
    begin
      while not QQ.eof do
      begin
      	CoefUaUs := QQ.findField('GL_COEFCONVQTE').AsFloat;
        Qte := QQ.findField('QTERECEP').AsFloat; // EN UA
        // Passage en UV
        if CoefUAUS <> 0 then
        begin
          QteRecep := QteRecep + (Qte * (COEFUAUS / FUV)) ;
        end else
        begin
          FUA := RatioMesure('PIE', QQ.findField('GL_QUALIFQTEACH').asString);
          QteRecep := QteRecep + (Qte * (FUA / FUV)) ;
        end;
        QQ.Next;
      end;
    end;
  FINALLY
    Ferme (QQ);
  END;

  if QteRecep >= TOBL.GetValue('GL_QTEFACT') then exit;
  // --
  TOBF := nil;
  if ModifOk then
  begin
//    TOBF := TOBDest.FindFirst (['GL_ARTICLE'],[TOBL.GetValue('GL_ARTICLE')],true); : correction FS#1702 pas de regroupement des articles divers
    TOBF := TOBDest.FindFirst (['GL_ARTICLE','GL_LIBELLE'],[TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_LIBELLE')],true);
  end;

  if TOBF = nil then
  begin
    inc(NumLigRea);
    TOBNL := TOB.Create ('LIGNE',TOBDest,-1);
    TOBNL.Dupliquer (TOBL,false,true);
    if (depdeb = '') and (depfin = '' ) then TOBNL.putvalue ('GL_DEPOT','');
    Qte := TOBNL.GetValue('GL_QTERESTE') * FUV / FUS;
    TOBL.PutValue('GL_QTEFACT',Qte);
    TOBL.PutValue('GL_QTESTOCK',Qte);
    TOBL.PutValue('GL_QTERESTE',Qte);
    TOBL.PutValue('GL_PUHT',(TOBL.GetValue('GL_DPA') / FUV) * FUs);
{$IFDEF BTP}
    AjouteLigneLien (TOBL,TOBNL,NumLigRea);
{$ENDIF}
  end else
  begin
  	if TOBL.GetValue('GL_AFFAIRE')<>TOBF.getValue('GL_AFFAIRE') then
    begin
    	// cas ou plusieurs lignes en provenance de +ieurs chantiers
      TOBF.putValue('GL_AFFAIRE','');
      TOBF.putValue('GL_AFFAIRE1','');
      TOBF.putValue('GL_AFFAIRE2','');
      TOBF.putValue('GL_AFFAIRE3','');
      TOBF.putValue('GL_AVENANT','');
    end;

    // PAssage UV en US
    Qte := TOBF.GetValue('GL_QTEFACT') * FUV / FUS;
    TOBF.PutValue('GL_QTEFACT',TOBL.GetValue('GL_QTEFACT')+Qte);
    TOBF.PutValue('GL_QTESTOCK',TOBL.GetValue('GL_QTEFACT')+Qte);
    TOBF.PutValue('GL_QTERESTE',TOBL.GetValue('GL_QTEFACT')+Qte);
    // Passage en US
{$IFDEF BTP}
    AjouteLigneLien (TOBL,TOBF,TOBF.GetValue('LIENLIGNE'));
{$ENDIF}
  end;
end;

function GestionStockIdentique (TOBL : TOB) : boolean;
var QQ : TQuery;
begin
  result := false;
  QQ := OpenSql ('SELECT GA_TENUESTOCK FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.getValue('GL_ARTICLE')+'"',True,-1,'',true);
  if not QQ.eof then
  begin
    if (QQ.findField('GA_TENUESTOCK').asString = TOBL.getValue('GL_TENUESTOCK')) then result := true;
  end;
  ferme (QQ);
end;

procedure FaisLeTriStd (TOBEntree : TOB);
  function FournPrinc (Article : string) : string;
  var QQ : Tquery;
  begin
    result := '';
    QQ := OpenSql ('SELECT GA_FOURNPRINC FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,-1,'',true);
    if not QQ.eof then Result := QQ.findfield('GA_FOURNPRINC').AsString;
    ferme (QQ);
  end;
var indice,NumLigRea : integer;
    TOBL : TOb;
    Fournisseur,Article : string;
//    mtpa : double;
begin
  NumLigRea := 0;
  for Indice :=0 to TOBEntree.detail.count -1 do
  begin
    TOBL := TOBEntree.detail[Indice];
    Fournisseur := TOBL.GetValue('GL_FOURNISSEUR');
    Article :=  TOBL.GetValue('GL_ARTICLE');
{$IFDEF BTP}
    if IsPrestationInterne (Article) then continue;
//    MtPa := RecupTarifAch (Fournisseur,Article,TurVente,True,False);
{$ENDIF}
    if (Fournisseur <> '') then
    begin
      if (FournPrinc(Article)<>Fournisseur)
{$IFDEF BTP}
          or (not GestionStockIdentique (TOBL))
          {or (TOBL.GetValue('GL_DPA') <> MtPa)} // pour l'instant on a pas de flag indiquant la modif
                                                 // du pa sur le document
{$ENDIF}
      then
      begin
        AjouteCumule (TOBDetailException,NumLigRea,TOBL,false);
      end else
      begin
        AjouteCumule (TOBDetailStd,NumLigRea,TOBL);
      end;
    end else
    begin
    	IF (WithStocks) and (not GestionStockIdentique (TOBL)) THEN
      begin
        AjouteCumule (TOBDetailException,NumLigRea,TOBL,false);
      end else
      begin
      	AjouteCumule (TOBDetailStd,NumLigRea,TOBL);
      end;
    end;
  end;
end;

function LanceTraitementFromDocs (var TOBPiece : TOB) : boolean;
var GetOut : boolean;
{$IFDEF BTP}
    Parametres : R_ParamTrait;
{$ENDIF}
begin
  Result := true;
  AvecHistoriqueConso := true;
  GetOut := false;
  TOBListeDocuments := TOB.Create ('LES DOCUMENTS A TRAITER',nil,-1);
  TOBDetailException := TOB.Create ('LIGNES EXCEPTIONS',nil,-1);
  TOBDetailStd := TOB.Create ('LIGNES STANDARD',nil,-1);
  TRY
{$IFDEF BTP}
    if ISPrepaLivFromAppro then
    begin
      parametres.ArtDeb := Artdeb;
      parametres.ArtFin := Artfin;
      parametres.DepDeb := Depdeb;
      parametres.DepFin := Depfin;
      parametres.Fam1 := Fam1;
      parametres.Fam2 := Fam2;
      parametres.Fam3 := Fam3;
      parametres.Fam1_ := Fam1_;
      parametres.Fam2_ := Fam2_;
      parametres.Fam3_ := Fam3_;
      parametres.DatDeb := DatDeb;
      parametres.DatFin := Datfin;
      GenereCommandeChantierFromRea (parametres);
    end;
{$ENDIF}
    DefiniListeDocuments (TOBListeDocuments,true);

    if AvecSelectionDocument then
      begin
      if not SelectionDocuments (GetNaturePieceCde,DatDeb,DatFin,TOBListeDocuments) then GetOut := true;
      end;
    if not GetOut then
      begin
      if RechercheDocuments (TOBListeDocuments) then
        begin
           result := TraitDocuments (TOBPiece);
        end else
        begin
          AuMoinsUn := false;
          result := false;
        end;
      end;
  FINALLY
    TOBListeDocuments.free;
    TOBDetailException.free;
    TOBDetailStd.free;
  END;
end;

procedure EnleveLesStandards(TOBLIG : TOB);
var TOBE,TOBR : TOB;
    Indice : integer;
begin
  if TOBDetailStd = nil then exit;
  if TOBDetailStd.detail.count = 0 then exit;
  for Indice := 0 to TOBDetailStd.detail.count -1 do
  begin
    TOBE := TOBDetailStd.detail[Indice];
    TOBR := TobLig.findFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG','GL_NUMLIGNE'],
                             [TOBE.GetValue('GL_NATUREPIECEG'),
                              TOBE.GetValue('GL_SOUCHE'),
                              TOBE.GetValue('GL_NUMERO'),
                              TOBE.GetValue('GL_INDICEG'),
                              TOBE.GetValue('GL_NUMLIGNE')],true);
    if TOBR <> nil then TOBR.free;
  end;
end;

procedure EnleveLesExceptions(TobLig : tob);
var TOBE,TOBR : TOB;
    Indice : integer;
begin
  if TOBDetailException = nil then exit;
  if TOBDetailException.detail.count = 0 then exit;
  for Indice := 0 to TOBDetailException.detail.count -1 do
  begin
    TOBE := TOBDetailException.detail[Indice];
    TOBR := TOBLig.findFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG','GL_NUMLIGNE'],
                             [TOBE.GetValue('GL_NATUREPIECEG'),
                              TOBE.GetValue('GL_SOUCHE'),
                              TOBE.GetValue('GL_NUMERO'),
                              TOBE.GetValue('GL_INDICEG'),
                              TOBE.GetValue('GL_NUMLIGNE')],true);
    if TOBR <> nil then TOBR.free;
  end;
end;

procedure EnlevelesAutresExceptions (TOBLTraite,TOBLIG : TOB);
var Indice : integer;
    TOBL: TOB;
    ArticleWanted : String;
begin
  if TOBLig = nil then exit;
  if TOBLig.detail.count = 0 then exit;
  ArticleWanted := TOBLTraite.GetValue('GL_ARTICLE');
  Indice := 0;
  repeat
    TOBL := TOBLig.detail[Indice];
    if (TOBL.GetValue('GL_ARTICLE')=ArticleWanted) and
        ((TOBL.GetValue('GL_NATUREPIECEG') <> TOBLTraite.getValue('GL_NATUREPIECEG')) OR
        (TOBL.GetValue('GL_SOUCHE') <> TOBLTraite.getValue('GL_SOUCHE')) OR
        (TOBL.GetValue('GL_NUMERO') <> TOBLTraite.getValue('GL_NUMERO')) OR
        (TOBL.GetValue('GL_INDICEG') <> TOBLTraite.getValue('GL_INDICEG')) OR
        (TOBL.GetValue('GL_NUMLIGNE') <> TOBLTraite.getValue('GL_NUMLIGNE'))) then
        TOBL.free else inc(Indice);
  until Indice > TOBLig.detail.count -1;
end;

end.
