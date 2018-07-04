{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q-,R-,S-,T-,U-,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 14/02/2000
Modifié le ... :   /  /
Description .. : TOF pour l'écran de génération de listes d'inventaire GCLISTEINV
Mots clefs ... : INVENTAIRE;STOCK
*****************************************************************}
unit UTofListeInv;

interface
uses Classes, Forms,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HRichOLE, UTOB, UTOF, EntGC, UtilGC, Vierge,HCtrls;

type
    THeader = class
      CodeListe : String;
      Depot : String;
      Fournisseur : String;
      NbLignes : Integer;
    end;

    // Cette classe conserve la liste des en-têtes générés ; elle tient compte
    //   des cas où l'on veut générer des listes distinctes par fournisseur ou non
    THeaderList = class(TList)
    private
      FSplitFourn : Boolean;
      FNextCode : Integer;
      FRacineCode : String;
      function GetHeader(index : integer) : THeader;

    public
      constructor Create(SplitFourn : Boolean; RacineCode : String);
      destructor Destroy; override;

      function GetListCode(Depot, Fournisseur : String) : String;
      function AddHeader(Depot, Fournisseur : String) : THeader;
      procedure IncNbLignes(Code : String);

      property Header[index : integer] : THeader read GetHeader;

    end;

    TOF_ListeInv = class(TOF)
    private
      FFlag : Boolean;
      FStockClos : Boolean;       // Inventaire à partir d'un stock clôturé (Arrêté de période)
      FReportMemo : THRichEditOLE;
      FHList : THeaderList;
      FCriTOB : TOB;
      FBigMother : TOB;
      TOBArtNonValid : TOB;
      FQ : TQuery;
      FDepot : String;
      DEPOT,DEPOTF : THVAlComboBox;
      function CheckGoodFields : boolean;
      function CheckBonSens : boolean;

      function GoodCondition(FieldEcran, FieldTable : String) : String;
      function AvecLesAND(Conditions : Array of String; SansLe1erAnd : Boolean) : String;
      function GetCodeListe(Depot : String) : String;

      procedure GenererListeInv;
      procedure GenereEntete(H : THeader);
      function GenereLigne : boolean;
      procedure GenereLots(LigneMere : TOB);

      procedure DoReport;
      procedure GereStockClos(FinMois, WithDate : boolean);
{$IFDEF BTP}
      procedure DEPOTChange(Sender: TObject);
{$ENDIF}
    published
      procedure OnLoad; override;
      procedure OnUpdate; override;
      procedure OnArgument (S : String ) ; override ;

      procedure OnExitDate(Sender: TObject);

    end;

//const DefaultLot = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';


function VideToX(S,X : String) : String;
Function RatioMesure(Cat, Mesure : String ) : Double;
function Ratioize(P, UStock, UVente, PPQ : Double) : Double;

implementation
uses Controls, StdCtrls, SysUtils, M3FP, HMsgBox, HEnt1, HStatus, InvUtil;


// Noms des champs critères de sélection (les contrôles portent les mêmes noms sur la fiche, ainsi que les mêmes avec un underscore en plus)
const FieldNamez : Array[1..12] of String = ('GA_CODEARTICLE','GA_FAMILLENIV1','GA_FAMILLENIV2','GA_FAMILLENIV3','GA_LIBREART1','GA_LIBREART2','GA_LIBREART3','GA_COLLECTION','GAT_REFTIERS','GA_FOURNPRINC','GQ_DEPOT','GQ_EMPLACEMENT');


// Initialisations au chargement de la fiche
procedure TOF_ListeInv.OnLoad;

var i : integer;
    Ctl : TControl ;
begin
inherited;
(Ecran.FindComponent('INTITULE') as TWinControl).SetFocus;
GereStockClos(False,False);

with (Ecran.FindComponent('DATEINV') as THEdit) do
  begin
  OnExit := OnExitDate;
  if FStockClos = True
   then Text := DateToStr(DebutDeMois(V_PGI.DateEntree)-1)
   else Text := DateToStr(V_PGI.DateEntree);
  end;

//THValComboBox(Ecran.FindComponent(' METTRE INV DANS LE COMBO
SetControlText('NATUREPIECEG', 'INV');

// Récupération du mémo de rapport
with Ecran do
 for i := 0 to ComponentCount-1 do
  if Components[i] is THRichEditOLE then
    FReportMemo := (Components[i] as THRichEditOLE);

// Libellés familles
for i := 1 to 3 do
 THLabel(Ecran.FindComponent('TGA_FAMILLENIV'+inttostr(i))).Caption := RechDom('GCLIBFAMILLE', 'LF'+inttostr(i), false) ;

// Libellés Table libre Article
GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 3, '');

// Libellé Dépôt ou Etablissement
if not VH_GC.GCMultiDepots then
   begin
    Ctl := GetControl('TGQ_DEPOT') ;
    if (Ctl <> Nil) and (Ctl is THLabel) then THLabel(Ctl).Caption := 'Etablissement' ;
   end ;
SetControlText('GQ_DEPOT', VH_GC.GCDepotDefaut);
SetControlText('GQ_DEPOT_', VH_GC.GCDepotDefaut);
{$IFDEF BTP}
DEPOTChange(Self);
{$ENDIF}
SetControlProperty('GA_CODEARTICLE','Plus',' AND ((GA_TYPEARTICLE="MAR") OR (GA_TYPEARTICLE="ARP"))');
SetControlProperty('GA_CODEARTICLE_','Plus',' AND ((GA_TYPEARTICLE="MAR") OR (GA_TYPEARTICLE="ARP"))');
end;

// Lancer la génération au click sur la mouette
procedure TOF_ListeInv.OnUpdate;
var IOErr : TIOErr;
begin
inherited;
if CheckBonSens then
  begin
  if not FStockClos then
    if PGIAsk('Confirmez-vous la création de la liste d''inventaire ?') <> mrYes then Exit;
  FBigMother := Nil;
  TOBArtNonValid := Nil;
  IOErr := Transactions(GenererListeInv, 0);
   case IOErr of
    oeUnknown : PGIBox('Erreur à la génération de listes !', Ecran.Caption);
    oeSaisie : PGIBox('Erreur à la génération de listes !'#13'(en cours de traitement par un autre utilisateur)', Ecran.Caption);
   end;
  if FBigMother<>Nil then FBigMother.free;
  if TOBArtNonValid<>Nil then TOBArtNonValid.free;
  end;
end;

// Vérifie que les critères sont bien dans le sens alphanumérique
function TOF_ListeInv.CheckBonSens : boolean;

var i : integer;
    s1,s2 : String;
    Ctl : TControl ;
begin
result := false;
if FFlag then exit;

for i := Low(FieldNamez) to High(FieldNamez) do
  begin
  Ctl := GetControl(FieldNamez[i]) ;
  if (Ctl <> Nil) then
    begin
    if (Ctl is THEdit) then
      begin
      s1 := (GetControl(FieldNamez[i]) as THEdit).Text;
      s2 := (GetControl(FieldNamez[i]+'_') as THEdit).Text;
      end else
      begin
      s1 := (GetControl(FieldNamez[i]) as THValComboBox).Value;
      s2 := (GetControl(FieldNamez[i]+'_') as THValComboBox).Value;
      end;

    if (s1 <> '') and (s2 <> '') and (s1 > s2) then
      begin
      PGIInfo('Incohérence de critère : '+(Ecran.FindComponent('T'+FieldNamez[i]) as THLabel).Caption, Ecran.Caption);
      FFlag := true;
      (Ecran.FindComponent(FieldNamez[i]) as TWinControl).SetFocus;
      FFlag := false;
      exit;
      end;
    end;
  end;

result := true;
end;

procedure TOF_ListeInv.OnExitDate(Sender: TObject);
begin
//with (Ecran.FindComponent('DATEINV') as THEdit) do
//  Text := DateToStr(FinDeMois(StrToDate(Text)));
end;

// ------------- CLASSE THEADERLIST --------------------------------------------
function THeaderList.GetHeader(index : integer) : THeader;
begin
result := Items[index];
end;

// Dire à la création si on veut des listes distinctes par fournisseur (true) ou non (false)
constructor THeaderList.Create(SplitFourn : Boolean; RacineCode : String);
begin
inherited Create;
FSplitFourn := SplitFourn;
FNextCode := 1;
FRacineCode := RacineCode;
while ExisteSQL('SELECT GIE_CODELISTE FROM LISTEINVENT WHERE GIE_CODELISTE="' + FRacineCode + Format('%.6d',[FNextCode]) + '"') do inc(FNextCode);
end;

destructor THeaderList.Destroy;
var i : integer;
begin
for i := 0 to Count-1 do
 Header[i].Free;
inherited;
end;

// Retourne le code liste correspondant au dépot/[fournisseur] ; retourne vide si inexistante
function THeaderList.GetListCode(Depot, Fournisseur : String) : String;

var i : integer;
begin
result := '';
for i := 0 to Count-1 do
  if Header[i].Depot = Depot then
    if (not FSplitFourn) or (Header[i].Fournisseur = Fournisseur) then
      begin
      result := Header[i].CodeListe;
      break;
      end;
end;

// Ajoute un en-tête à la liste et le retourne (choix du code automatique)
function THeaderList.AddHeader(Depot, Fournisseur : String) : THeader;
begin
result := THeader.Create;
result.Depot := Depot;
result.Fournisseur := Fournisseur;
result.CodeListe := FRacineCode + Format('%.6d',[FNextCode]);
result.NbLignes := 0;
Add(result);
Inc(FNextCode);
end;

procedure THeaderList.IncNbLignes(Code : String);
var i : integer;
begin
for i := 0 to Count-1 do
 if Header[i].CodeListe = Code then
   begin
   Inc(Header[i].NbLignes);
   break;
   end;
end;
// -----------------------------------------------------------------------------


// Génère un en-tête de liste d'inventaire
procedure TOF_ListeInv.GenereEntete(H : THeader);

begin
  with TOB.Create('LISTEINVENT', FBigMother, -1) do
  begin
    PutValue('GIE_CODELISTE', H.CodeListe);
    PutValue('GIE_LIBELLE', Copy(FCriTOB.GetValue('INTITULE'),1,35));
    PutValue('GIE_DEPOT', H.Depot);
    PutValue('GIE_DATEINVENTAIRE', FCriTOB.GetValue('DATEINV'));
    PutValue('GIE_DATECREATION', NowH);
    PutValue('GIE_DATEMODIF', NowH);
    PutValue('GIE_CREATEUR', V_PGI.USER);
    PutValue('GIE_UTILISATEUR', V_PGI.USER);
    PutValue('GIE_VALIDATION', '-');
    PutValue('GIE_STOCKCLOS', FCriTOB.GetValue('STOCKCLOS'));
    PutValue('GIE_NATUREPIECEG', FCriTOB.GetValue('NATUREPIECEG'));
  end;
end;

// Retourne le code liste correspondant au dépot/[fournisseur]
//  Si la liste n'existe pas, l'en-tête est généré
function TOF_ListeInv.GetCodeListe(Depot : String) : String;

var Fourn : String;
    H : THeader;
begin
if (ctxMode in V_PGI.PGIContexte)
 then Fourn := FQ.FindField('GA_FOURNPRINC').AsString
 else Fourn := FQ.FindField('GAT_REFTIERS').AsString;

result := FHList.GetListCode(Depot, Fourn);
if result = '' then  // Si l'en-tête n'a pas encore été créé...
  begin
  H := FHList.AddHeader(Depot,Fourn);  // ...on l'ajoute à la liste
  GenereEntete(H);   // et on le génère
  result := H.CodeListe;
  end;
end;


function VideToX(S,X : String) : String;
begin
if S = '' then result := X
          else result := S;
end;

Function RatioMesure(Cat, Mesure : String) : Double;
var TOBM : TOB;
    X : Double;
begin
TOBM := VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'], [Cat, Mesure], False);
X := 0; if TOBM <> nil then X := TOBM.GetValue('GME_QUOTITE'); if X = 0 then X := 1.0;
result := X;
end;

// Remet P en unité de stock
function Ratioize(P, UStock, UVente, PPQ : Double) : Double;
begin
result := (P * UStock)/(PPQ * UVente);
end;

// Génère une ligne de liste d'inventaire
function TOF_ListeInv.GenereLigne : boolean;

var CodeListe   : String;
    GaArticle   : String;
    GereParLot  : String;
    Depot       : String;
    FUS         : Double;
    FUV         : Double;
    FPPQ        : Double;
    FEMPL : String;
    FPHY        : Double;
    FDPA        : Double;
    FPMAP       : Double;
    FDPR        : Double;
    FPMRP       : Double;
    LM          : TOB;
    ListeMere   : TOB;
    TobTest     : TOB;
    QP : TQtePrixRec;
begin
  result := false;

  GaArticle := FQ.FindField('GA_ARTICLE').AsString;
  Depot := FDepot;

  // Vérifier que l'article n'est pas déjà inséré dans une liste non validée :
  if (TOBArtNonValid.FindFirst(['GIL_ARTICLE'], [GaArticle], true) <> nil) then exit;

  CodeListe := GetCodeListe(Depot);
  ListeMere := FBigMother.FindFirst(['GIE_CODELISTE'], [CodeListe], false);
  if ListeMere = nil then raise Exception.Create('Liste parente non trouvée');

  //FV1 : 05/01/2014 - FS#1366 - COFRAMENAL - message d'erreur à la création d'une liste d'inventaire
  TOBTEST := ListeMere.FindFirst(['GIL_ARTICLE','GIL_DEPOT'],[GaArticle, Depot], False);
  if TOBTEST <> nil then
  begin
    result := False;
    Exit;
  end;

  FHList.IncNbLignes(CodeListe);

{MRNONGESTIONLOT}
  GereParLot := FQ.FindField('GA_LOT').AsString;  
  FUS := RatioMesure('PIE', FQ.FindField('GA_QUALIFUNITESTO').AsString);
  FUV := RatioMesure('PIE', FQ.FindField('GA_QUALIFUNITEVTE').AsString); if FUV=0 then FUV:=1;
  // FPPQ := FQ.FindField('GA_PRIXPOURQTE').AsFloat; if FPPQ=0 then FPPQ:=1;  MR : FAUX, GAPRIXPOURQTE ne concerne que les prix de Vente
  FPPQ:=1;
  FEMPL := FQ.FindField('GQ_EMPLACEMENT').AsString;
  FPHY  := FQ.FindField('GQ_PHYSIQUE').AsFloat;
  FDPA  := Ratioize(FQ.FindField('GQ_DPA').AsFloat, FUS, FUV, FPPQ);
  FPMAP := Ratioize(FQ.FindField('GQ_PMAP').AsFloat, FUS, FUV, FPPQ);
  FDPR  := Ratioize(FQ.FindField('GQ_DPR').AsFloat, FUS, FUV, FPPQ);
  FPMRP := Ratioize(FQ.FindField('GQ_PMRP').AsFloat, FUS, FUV, FPPQ);

  if FCriTOB.GetValue('STOCKCLOS')='X' then
     begin
     // Recalcul de la quantité en stock, en partant du dernier arrêté de période
     QP := GetQtePrixDateListe(GaArticle, Depot, FCriTOB.GetValue('DATEINV'));
     if not QP.SomethingReturned then QP.Qte := 0;
     end else
     begin
     // Reprise du stock courant dans la table DISPO
     QP.Qte := FPHY;
     end;

  LM := TOB.Create('LISTEINVLIG', ListeMere, -1);
  with LM do
  begin
    PutValue('GIL_CODELISTE', CodeListe);
    PutValue('GIL_ARTICLE', GaArticle);
    PutValue('GIL_CODEARTICLE', FQ.FindField('GA_CODEARTICLE').AsString);
    PutValue('GIL_LOT', GereParLot);
    PutValue('GIL_DEPOT', Depot);
    PutValue('GIL_DATESAISIE', StrTodate(ThEdit(getControl('DATEINV')).Text));
    PutValue('GIL_SAISIINV', '-');
    PutValue('GIL_INVENTAIRE', 0);
    PutValue('GIL_QTEPHOTOINV', QP.Qte);
    PutValue('GIL_EMPLACEMENT', FEMPL);
    PutValue('GIL_DPA', FDPA);
    PutValue('GIL_PMAP', FPMAP);
    PutValue('GIL_DPR', FDPR);
    PutValue('GIL_PMRP', FPMRP);
    PutValue('GIL_DPAART', Ratioize(FQ.FindField('GA_DPA').AsFloat, FUS, FUV, FPPQ));
    PutValue('GIL_PMAPART', Ratioize(FQ.FindField('GA_PMAP').AsFloat, FUS, FUV, FPPQ));
    PutValue('GIL_DPRART', Ratioize(FQ.FindField('GA_DPR').AsFloat, FUS, FUV, FPPQ));
    PutValue('GIL_PMRPART', Ratioize(FQ.FindField('GA_PMRP').AsFloat, FUS, FUV, FPPQ));
    PutValue('GIL_DPASAIS', 0);
    PutValue('GIL_PMAPSAIS', 0);
    PutValue('GIL_DPRSAIS', 0);
    PutValue('GIL_PMRPSAIS', 0);
  end;

  if GereParLot = 'X' then GenereLots(LM);
  result := true;

end;

// Génère les lots d'un article géré par lots
procedure TOF_ListeInv.GenereLots(LigneMere : TOB);

var CodeListe, GaArticle : String;
    QLot : TQuery;
begin
  CodeListe := LigneMere.GetValue('GIL_CODELISTE');
  GaArticle := LigneMere.GetValue('GIL_ARTICLE');

  QLot := OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE="'+GaArticle+'" AND GQL_DEPOT="'+LigneMere.GetValue('GIL_DEPOT')+'"', true);

  while not QLot.EOF do
  begin
    with TOB.Create('LISTEINVLOT', LigneMere, -1) do
    begin
      PutValue('GLI_CODELISTE', CodeListe);
      PutValue('GLI_ARTICLE', GaArticle);
      PutValue('GLI_NUMEROLOT', QLot.FindField('GQL_NUMEROLOT').AsString);
      PutValue('GLI_SAISIINV', '-');
      PutValue('GLI_INVENTAIRE', 0);
      PutValue('GLI_QTEPHOTOINV', QLot.FindField('GQL_PHYSIQUE').AsFloat);
    end;
    QLot.Next;
  end;
  Ferme(QLot);
end;

// Retourne une condition SQL correcte en fonction des champs renseignés ou non
function TOF_ListeInv.GoodCondition(FieldEcran, FieldTable : String) : String;

var v1, v2 : String;
begin
v1 := FCriTOB.GetValue(FieldEcran);
v2 := FCriTOB.GetValue(FieldEcran+'_');

if v2 = '' then
  begin
  if v1 = '' then result := ''
             else result := FieldTable + '>="' + v1 + '"';
  end else
  begin
  if v1 = '' then result := FieldTable + '<="' + v2 + '"'
             else result := FieldTable + '>="' + v1 + '" AND ' + FieldTable + '<="' + v2 + '"';
  end;
end;

// Concatène un tableau de conditions SQL en éliminant les chaines vides et en mettant AND entre chaque condition
function TOF_ListeInv.AvecLesAND(Conditions : Array of String; SansLe1erAnd : Boolean) : String;

var i : integer;
    first : boolean;
begin
     first := SansLe1erAnd;
     result := '';
     for i := Low(Conditions) to High(Conditions) do
       if Conditions[i] <> '' then
         begin
         if not first then result := result + ' AND '
                      else first := false;
         result := result + Conditions[i];
         end;
end;

// Vérifie que tous les champs à renseigner le sont
function TOF_ListeInv.CheckGoodFields : boolean;

begin
result := false;

if FCriTOB.GetValue('INTITULE') = '' then
  begin PGIInfo('Vous devez renseigner l''intitulé de liste', Ecran.Caption);
        (Ecran.FindComponent('INTITULE') as TWinControl).SetFocus;
        exit; end;
if FCriTOB.GetValue('RACINECODE') = '' then
  begin PGIInfo('Vous devez renseigner le préfixe du code liste', Ecran.Caption);
        (Ecran.FindComponent('RACINECODE') as TWinControl).SetFocus;
        exit; end;

if FCriTOB.GetValue('STOCKCLOS')='X' then
  begin
  if String(FCriTOB.GetValue('DATEINV')) = '' then
    begin PGIInfo('Vous devez renseigner la date d''inventaire', Ecran.Caption);
          (Ecran.FindComponent('DATEINV') as TWinControl).SetFocus;
          exit; end;
//  if StrToDate(FCriTOB.GetValue('DATEINV')) >= Date then
  if StrToDate(FCriTOB.GetValue('DATEINV')) >= V_PGI.DateEntree then
    begin PGIInfo('La date d''inventaire doit être antérieure à la date de connexion', Ecran.Caption);
          (Ecran.FindComponent('DATEINV') as TWinControl).SetFocus;
          exit; end;
  end;

if FCriTOB.GetValue('NATUREPIECEG') = '' then
  begin PGIInfo('Vous devez renseigner la nature de pièce', Ecran.Caption);
        (Ecran.FindComponent('NATUREPIECEG') as TWinControl).SetFocus;
        exit; end;

result := true;
end;

procedure TOF_ListeInv.GereStockClos(FinMois, WithDate : boolean);
begin
FStockClos := FinMois;

//SetControlEnabled('DATEINV', FStockClos);
SetControlChecked('CHKSPLITFOURN', FStockClos);

if WithDate then
  begin
  with (Ecran.FindComponent('DATEINV') as THEdit) do
    begin
    if FStockClos = True then Text := DateToStr(FinDeMois(StrToDate(Text)))
    else Text := DateToStr(V_PGI.DateEntree);
    end;
  end;
end;

// Affiche le rapport de fin de génération
procedure TOF_ListeInv.DoReport;

var i : integer;
    blancs, s, t : string;
    BlocNote : TStringList;

 function DeA(Field : String) : String;
 var f,t : string;
     Ctl : TControl ;
 begin
 result := '';
 Ctl := GetControl(Field) ;
 if (Ctl <> Nil) then
   begin
   if Ctl is THValComboBox then
     begin
     f := '"'+THValComboBox(Ecran.FindComponent(Field)).Text+'"';
     t := '"'+THValComboBox(Ecran.FindComponent(Field+'_')).Text+'"';
     end else
     begin
     f := THCritMaskEdit(Ecran.FindComponent(Field)).Text;//FCriTOB.GetValue(Field); tcombobox
     t := THCritMaskEdit(Ecran.FindComponent(Field+'_')).Text;//FCriTOB.GetValue(Field+'_');
     end;
   result := 'de '+f+' à '+t;
   if FCriTOB.GetValue(Field) = '' then result := 'jusqu''à '+t;
   if FCriTOB.GetValue(Field+'_') = '' then result := 'à partir de '+f;
   end;
 end;

 procedure AddFormattedLine(var St : string; Field : String);
 begin
 if (FCriTOB.GetValue(Field) <> '') or (FCriTOB.GetValue(Field+'_') <> '')
  then St := St + '    \plain\f5\fs16\ul '+THLabel(Ecran.FindComponent('T'+Field)).Caption+' :\plain\f5\fs16  '+DeA(Field)+' \par';
 end;

begin
BlocNote := TStringList.Create;
with FReportMemo.Lines do
  begin
  if FHList.Count <= 0 then exit;
  Add('------------------------------------');
  Add(' Liste(s) générée(s) :');
  Add('------------------------------------');
  blancs := ''; for i := 1 to Length(FHList.Header[0].CodeListe)+3 do blancs := blancs + ' ';
  for i := 0 to FHList.Count-1 do
    begin
    Add(FHList.Header[i].CodeListe+' : '+RechDom('GCDEPOT',FHList.Header[i].Depot,false));
    Add(blancs+IntToStr(FHList.Header[i].NbLignes)+' article(s)');

    with BlocNote do
      begin
      Clear;
      t := '{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss Tahoma;}{\f4\fswiss Arial;}{\f5\fswiss\fcharset1 Arial;}}'+
          '{\colortbl\red0\green0\blue0;\red255\green0\blue0;}'+
          '\deflang1036\pard\plain\f5\fs16\b Crit\''e8res de génération :\plain\f5\fs16  \par';

      AddFormattedLine(t, 'GA_CODEARTICLE');
      AddFormattedLine(t, 'GA_FAMILLENIV1');
      AddFormattedLine(t, 'GA_FAMILLENIV2');
      AddFormattedLine(t, 'GA_FAMILLENIV3');
      AddFormattedLine(t, 'GA_LIBREART1');
      AddFormattedLine(t, 'GA_LIBREART2');
      AddFormattedLine(t, 'GA_LIBREART3');

      if (ctxMode in V_PGI.PGIContexte) then
        begin
        AddFormattedLine(t, 'GA_COLLECTION');
        if FCriTOB.GetValue('CHKSPLITFOURN') <> 'X'
          then AddFormattedLine(t, 'GA_FOURNPRINC')
          else t := t + '    \plain\f5\fs16\ul '+THLabel(Ecran.FindComponent('TGA_FOURNPRINC')).Caption+' :\plain\f5\fs16  '+FHList.Header[i].Fournisseur+' \par';
        end else
        begin
        if FCriTOB.GetValue('CHKSPLITFOURN') <> 'X'
          then AddFormattedLine(t, 'GAT_REFTIERS')
          else t := t + '    \plain\f5\fs16\ul '+THLabel(Ecran.FindComponent('TGAT_REFTIERS')).Caption+' :\plain\f5\fs16  '+FHList.Header[i].Fournisseur+' \par';
        AddFormattedLine(t, 'GQ_EMPLACEMENT');
        end;
      if FHList.Header[i].NbLignes > 1 then s := 's' else s := '';
      t := t + '  \par \plain\f5\fs16\b '+IntToStr(FHList.Header[i].NbLignes)+' article'+s+' compilé'+s+'\plain\f5\fs16  \par }';
      Add(t);

      with TOB.Create('LISTEINVENT', nil, -1) do
        begin
        SelectDB('"'+FHList.Header[i].CodeListe+'"', nil);
        PutValue('GIE_BLOCNOTE', BlocNote.Text);
        UpdateDB;
        Free;
        end;

      end;
    end;
  end;
BlocNote.Free;
end;

// Séléctionne les articles en fonction des critères et lance la génération des listes
procedure TOF_ListeInv.GenererListeInv;

var RequeteSQL : String;  // Requête SQL générée par les critères
    SQLDepot   : String;  // Requête SQL retournant la liste des dépôts sélectionnés
    SQLArtNonValid  : string;
    //StockVide   : string;
    ListeDepot  : String;
    TOBDepot    : TOB;
    TOBD        : TOB;
    QA          : TQuery;
    QD          : TQuery ;
    i           : Integer;
    NbLig       : Integer;
    NbListe     : integer;
    H : THeader;
begin
FCriTOB := TOB.Create('criteres liste inventaire', nil, -1);

with FCriTOB do
  begin
  // Les noms des contrôles sur la fiche doivent être les mêmes que les champs
  //   auxquels ils correspondent
  for i := Low(FieldNamez) to High(FieldNamez) do
    begin
    AddChampSup(FieldNamez[i], false);
    AddChampSup(FieldNamez[i]+'_', false);
    PutValue(FieldNamez[i], '');
    PutValue(FieldNamez[i]+'_', '');
    end;
  AddChampSup('CHKSPLITFOURN', false);
  AddChampSupValeur('CHKARTMOUV','X');     // Par défaut, seuls les articles mouvementés figureront dans la liste
  AddChampSupValeur('CHKDIFFZERO','X');
  AddChampSup('INTITULE', false);
  AddChampSup('RACINECODE', false);
  AddChampSup('DATEINV', false);
  AddChampSup('STOCKCLOS', false);
  AddChampSup('NATUREPIECEG', false);

  GetEcran(Ecran, nil);
  if FStockClos = True
   then PutValue('STOCKCLOS', 'X')
   else PutValue('STOCKCLOS', '-') ;
  end;

if CheckGoodFields then
  begin
  ListeDepot := GoodCondition('GQ_DEPOT', 'GIE_DEPOT');
  if ExisteSQL('SELECT GIE_CODELISTE FROM LISTEINVENT '+
  'WHERE GIE_VALIDATION<>"X" AND ' + ListeDepot) then
  begin
    if PGIAsk('Il existe déjà une autre liste non encore validée pour ce même dépôt, Voulez-vous continuer ?') <> MrYes then
    begin
      FCriTOB.Free;
      Exit;
    end;
    if FStockClos then
      if PGIAsk('La création d''une liste d''inventaire en fin de mois#13'+
              ' est un traitement long qui nécessite des clôtures de stock régulières.#13'+
              ' Voulez vous continuer ?') <> mrYes then Exit;
  end;

  // Chargement de la liste des dépôts sélectionnés
  TOBDepot := TOB.Create ('Liste-Depot', Nil, -1) ;
  ListeDepot := GoodCondition('GQ_DEPOT', 'GDE_DEPOT');
  SQLDepot := 'SELECT GDE_DEPOT from DEPOTS where '+ListeDepot;

  QD := OpenSQL(SQLDepot, True) ;
  if Not QD.EOF then TOBDepot.LoadDetailDB('','','',QD,False,True) ;
  Ferme (QD) ;

  NbLig := 0;
  NbListe := 0;

  // Pour chaque dépôt, génération d'une liste contenant :
  //   - soit uniquement les articles mouvementés (ayant une ligne dans la table DISPO)
  //   - soit tous les articles, même ceux n'ayant jamais été mouvementés (Contexte MODE uniquement).
  // Ce choix est fonction de la valeur du booléan CHKARTMOUV
  for i:=0 to TobDepot.Detail.Count-1 do
    begin
    TOBD := TobDepot.Detail[i] ;
    FDepot := TOBD.GetValue('GDE_DEPOT');


    FHList := THeaderList.Create(FCriTOB.GetValue('CHKSPLITFOURN') = 'X', FCriTOB.GetValue('RACINECODE'));
    FBigMother := TOB.Create('Big Mother', nil, -1);

    RequeteSQL := 'SELECT GA_ARTICLE, GA_CODEARTICLE, GA_LOT, GQ_DEPOT, GAT_REFTIERS, GQ_EMPLACEMENT, GQ_PHYSIQUE, '+
    							'GQ_DPA, GQ_PMAP, GQ_DPR, GQ_PMRP, GA_DPA, GA_PMAP, GA_DPR, GA_PMRP, GA_QUALIFUNITESTO, '+
                  'GA_QUALIFUNITEVTE, GA_PRIXPOURQTE '+
                  'FROM DISPO '+
                  'LEFT JOIN ARTICLE ON GA_ARTICLE = GQ_ARTICLE '+
                  'LEFT JOIN ARTICLETIERS ON GA_CODEARTICLE = GAT_ARTICLE '+
                  'WHERE GA_TENUESTOCK="X" AND GA_STATUTART<>"GEN" AND GQ_CLOTURE="-" AND GQ_DEPOT="'+FDepot+'" AND ' +
                  'GA_FERME<>"X" '+
									 AvecLesAND([GoodCondition('GA_CODEARTICLE','GA_CODEARTICLE'),
                               GoodCondition('GA_FAMILLENIV1','GA_FAMILLENIV1'),
                               GoodCondition('GA_FAMILLENIV2','GA_FAMILLENIV2'),
                               GoodCondition('GA_FAMILLENIV3','GA_FAMILLENIV3'),
                               GoodCondition('GA_LIBREART1','GA_LIBREART1'),
                               GoodCondition('GA_LIBREART2','GA_LIBREART2'),
                               GoodCondition('GA_LIBREART3','GA_LIBREART3'),
                               GoodCondition('GAT_REFTIERS','GAT_REFTIERS'),
                               GoodCondition('GQ_EMPLACEMENT','GQ_EMPLACEMENT') ], false);

    // Chargement des articles existant déjà dans une liste du dépôt non encore validée,
    // afin de pouvoir refuser parmi les articles sélectionnés, ceux figurant déjà dans cette liste
    TOBArtNonValid := TOB.Create ('Liste-Article', Nil, -1) ;
    SQLArtNonValid := 'SELECT GIL_ARTICLE FROM LISTEINVLIG LEFT JOIN LISTEINVENT ON GIL_CODELISTE=GIE_CODELISTE '+
                      'WHERE GIE_VALIDATION<>"X" AND GIE_DEPOT="'+FDepot+'"';
    QA := OpenSQL(SQLArtNonValid, True) ;
    if Not QA.EOF then TOBArtNonValid.LoadDetailDB('','','',QA,False,True) ;
    Ferme (QA) ;

    // Ouverture de la requête de génération des nouveaux articles en Liste
    FQ := OpenSQL(RequeteSQL, true);
    InitMove(100,'');

    if (FQ.EOF) and (FCriTOB.GetValue('CHKARTMOUV')='X') then
       begin
       // Si pas de liste déjà existante pour le dépôt, création d'une liste vide (sans ligne)
       // permettant d'intégrer par la suite un inventaire transmis par Terminal Portable.
       if Not(ExisteSQL('SELECT GIE_CODELISTE FROM LISTEINVENT WHERE GIE_DEPOT="'+FDepot+'"')) then
          begin
          H := FHList.AddHeader(FDepot,'');
          GenereEntete(H);
          end;
       end;

    while not FQ.EOF do
      begin
      // Génération d'une ligne par article sélectionné
      if GenereLigne then inc(NbLig);
      FQ.Next;
      MoveCur(false);
      end;
    Ferme(FQ);
    TOBArtNonValid.Free ;
    TOBArtNonValid := Nil;
    FiniMove;
    FBigMother.InsertDB(nil, true);
    DoReport;
    FBigMother.Free;
    FBigMother := Nil;
    NbListe := NbListe + FHList.Count;
    FHList.Free;
    end;

  TOBDepot.Free ;
  if NbListe = 0
    then PGIInfo('Aucune liste générée', Ecran.Caption)
    else PGIInfo(IntToStr(NbListe)+' liste(s) générée(s),'#13+IntToStr(NbLig)+' article(s) inséré(s)', Ecran.Caption);
  end;

FCriTOB.Free;
end;

Procedure AGLGCListeSaisie_StockClos( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_ListeInv(totof).GereStockClos(Parms[1],True);
end;

procedure TOF_ListeInv.OnArgument(S: String);
begin
  inherited;
{$IFDEF BTP}
  DEPOT := THVAlComboBox(GetCOntrol('GQ_DEPOT'));
  DEPOTF := THVAlComboBox(GetCOntrol('GQ_DEPOT_'));
  DEPOT.OnExit  := DepotChange;
  DEPOTF.OnExit := DepotChange;
{$ENDIF}
end;
{$IFDEF BTP}
procedure TOF_ListeInv.DEPOTChange(Sender: TObject);
var chainedep,chainefin : string;
begin
  if (Depot.Value  <> DepotF.value) or (Depot.value = '') or (DepotF.value = '' ) then
  begin
    if (Depot.value <> '') then ChaineDep := ' AND GEM_DEPOT>="'+ Depot.value + '"'
                           else ChaineDep := '';
    if (DepotF.value <> '') then Chainefin := ' AND GEM_DEPOT<="'+ DepotF.value + '"'
                            else Chainefin := '';
    SetControlProperty('GQ_EMPLACEMENT','Plus',ChaineDep+chaineFin );
  end else
  begin
    chainedep := ' AND GEM_DEPOT="'+Depot.value+'"';
    SetControlProperty('GQ_EMPLACEMENT','Plus',ChaineDep );
  end;
end;
{$ENDIF}
initialization
RegisterClasses([TOF_ListeInv]);
RegisterAglProc( 'GCListeSaisie_StockClos', True, 1, AGLGCListeSaisie_StockClos);

end.
