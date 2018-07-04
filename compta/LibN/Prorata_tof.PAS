{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 22/11/2001
Modifié le ... : 04/03/2003
Description .. : Source TOF de la TABLE : CPPRORATATVA ()
Suite ........ : Passage en eAGL
Mots clefs ... : TOF;CPPRORATATVA
*****************************************************************}
// Lors de la remise en place du champ PA_TOUSJAL ,
// décommenté les lignes avec *+* et supprimer les lignes avec *-*
Unit Prorata_tof ;

Interface

Uses StdCtrls,    // CbChecked, CbUnchecked
     Controls,    // MrNo
     Classes,     // registerclasses
{$IFDEF EAGLCLIENT}
     MaineAGL,    // AGLLanceFiche
     eMul,        // TFMul
{$ELSE}
     Fe_Main,     // AGLLanceFiche
     Mul,         // TFMul
     DB,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     HDB,
     SaisBor,     // LanceSaisieFolio
{$ENDIF}
     Saisie,      // TrouveEtLanceSaisie
     forms,       // TForm
     sysutils,    // DateToStr, FormatFloat, IntToStr, StrToDate, StrToInt, EncodeDate, Date
     ComCtrls,    // TTabSheet; TPageControl
     UTob,        // TOB
     HCtrls,      // THValComboBoxn, THEdit, THLapen, OpenSQL, Ferme, ReadTokenST, USDateTime
     HMsgBox,     // THMsgBox
     UTOF,        // TOF
     Htb97,       // TToolBarButton97
     Ent1,        // GetLibelleTableLibre, VH, HT2TVA, TVA2TAUX, TVA2CPTE, TVA2ENCAIS
     HEnt1,       // Format_String, EstSerie, String3, Arrondi, NumSemaine, V_PGI, TraduireMemoire, EnableControls, BeginTrans, CommitTrans, RollBack, TaConsult
     SaisUtil,    // Quelexo, GetNewNumJal, QuelExoDT
     SaisComm,    // TVA2Num
     ParamSoc,     // GetParamSocSecur
     UlibWindows  //TraductionTHMultiValComboBox //SG6 12/01/05 FQ 15155
     ;

Const
   Titre = 'Prorata de TVA' ;
   Dec2 = 2 ;

procedure CCLanceFiche_EcrProrataTVA(pszType : String);

Type
  TOF_CPPRORATATVA = Class (TOF)
    Exercice : THValComboBox ;
    JournalBis : THMultiValComboBox ;
    Tva      : THValComboBox ;
    Money    : THValComboBox ;
    ComboEtab: THValComboBox ;
    Datedeb : THEdit ;
    DateFin : THEdit ;
    DateGen : THEdit ;
    GenerationAu : THLabel ;
    General1 : THEdit ;
    General2 : THEdit ;
    Fiche : TFMul ;
    LEcr  : TStringList  ;
{$IFDEF EAGLCLIENT}
    ListeEcr : THGrid ;
    FListe   : THGrid ;
{$ELSE}
    ListeEcr : THDBGrid ;
    FListe   : THDBGrid ;
{$ENDIF}
    QlisteEcr : TQuery ;
    TabLib   : TTabSheet ;
    TobDev     : TOB ;
    TobEcr     : TOB ;
    TobProrata : TOB ;
    TobHT      : TOB ;
    TobTTC     : TOB ;
    TobTVA     : TOB ;
    TobGen     : TOB ;
    TobSec     : TOB ;
    Pages      : TPageControl ;
    MsgBox     : THMsgBox;
    procedure OnLoad                   ; override ;
    procedure OnClose                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    //Sg6 12/01/05 FQ 15155 Gestion du qualifpiece
    procedure qualifpiecechange(Sender : TObject);
    procedure InitMsgBox;
    procedure InitTablesLibres(TT: TTabSheet);
    procedure BOuvrirOnClick (Sender : TObject) ;
    procedure ChargeTobGeneraux;
    procedure AjouteLesAnas;
    procedure ChargeTobSection;
    procedure ExerciceOnChange(Sender: TObject);
    procedure ExerciceOnExit(Sender: TObject);
    Procedure ExerciceOnEnter (Sender : TObject);
    procedure ChargeTobDevise;
    procedure Traitement ;
    procedure ChargeTobProrata;
    function  RechercheCpteTVA(Ecr: Tob): String;
    procedure FListeDblClick(Sender: TObject);
    Function  RechercheTauxProrata (Ecr : TOB ; var QO : string; Var T1, T2 :Double) : Double ;
    function  RechercheTauxTVA(Ecr: TOB): Double;
    function  CalculMontantTVA (Ecr : Tob ; Lib : String ; Tx1, Tx2, TxTVA : Double ; Dec : Integer) : Double ;
    procedure Modiflignes(var TobTTC, TobHT, TobTVA : TOB ; TxP, TxTVA, MtTVA, MtTVAD : Double ; QO : String) ;
    Procedure NouvelLignes (var TobTTC, TobHT, TobTVA : TOB ; TxP, TxTVA, MtTVA, MtTVAD : Double ; QO : String) ;
    function  RechercheLigneCpteTVA(Compte: String): Integer;
    function  RechDecDev(Ecr: TOB): Integer;
    procedure MajCompte(Zone: String; E: Double; Exo: String; TobHT, TobTVA: Tob);
    procedure MajAnalytiq (var TobHT : TOB ; DecDev : Integer ; Zone, Exo : String) ;
    function  DupliqueLignes (var TobHT, TobTVA : TOB; pszAuxi : String) : Double ;
    function  RecalculHorsTaxes(Ecr : Tob ; Tx1, TxTVA : Double; Z: String; dec : Integer) : Double ;
    procedure ModifEcrAna(var TobE: Tob; Zone1, Zone2: String; NumPiece, Numligne : Integer ; Dat : TDateTime);
    function  RechercheDate(TobE: TOB): TDateTime;
    procedure LibereTobs;
    procedure MajZones(var TobE: TOB; DD: TdateTime; NumPiece, Numlig: Integer; pszAuxi : String );
    procedure Chargement ;
    function ChargeTobEcr (i : integer; var OkOk : Boolean ; var Averto : Boolean) : boolean;
    procedure MajTobs;
    function  AvecEcrModifiees : String;
    Function  RechercheLignettC : Integer ;
    procedure FaitLeWhereTablesLibres (Var S : String) ;
    function  VerifieLaPiece: boolean;
    function  VerifieTauxDefinitifSaisi : Boolean ;
    procedure JournalExit(Sender : TObject);
    function RetourneClauseIn(Chaine : string) : string;
    procedure MemoriseHTInitial(TobHT : Tob);
    procedure InitValid;
    {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
    {JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
    procedure ControlEtab;
    {JP 28/06/06 : FQ 16149 : pour appeler ControlEtab à la fin du chargement des filtres}
    procedure AfterSelectFiltre;
  end ;

Implementation

uses
    {$IFDEF MODENT1}
    CPProcMetier,
    {$ENDIF MODENT1}
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    ;

var QuelProrata : String ;
    EcrNonTraitee : Boolean ;
    AncienExo : String;

procedure CCLanceFiche_EcrProrataTVA(pszType : String);
begin
  AGLLanceFiche ('CP' , 'CPFICPRORATA', '', '' , pszType);
end;

procedure TOF_CPPRORATATVA.OnLoad ;
begin
  Inherited ;

  if AvecEcrModifiees='Sans' Then SetControlText('XX_WHERE1', '((E_QUALIFORIGINE<>"TV1" AND E_QUALIFORIGINE<>"TVI") Or (E_QUALIFORIGINE IS NULL))')  else
    if AvecEcrModifiees='Que' then SetControlText('XX_WHERE1', '(E_QUALIFORIGINE="TV1" OR E_QUALIFORIGINE="TVI")') else
      if AvecEcrModifiees='Indiff' Then SetControlText('XX_WHERE1', '') ;
end ;

procedure TOF_CPPRORATATVA.OnClose ;
begin
  LibereTobs ;
  LEcr.Clear ;
  LEcr.Free ;
  MsgBox.Free;
  Inherited ;
end ;

Procedure TOF_CPPRORATATVA.Chargement ;
var i : integer ;
    St : String ;
Begin
Lecr.Clear ;
if Fiche.FListe.AllSelected then
  begin
  While not Fiche.Q.Eof do
    begin
    St:=Format_String(Fiche.Q.FindField('E_EXERCICE').AsString,3)+';'
       +Fiche.Q.FindField('E_JOURNAL').AsString+';'
       +Format_String(Fiche.Q.FindField('E_QUALIFPIECE').AsString,1)+';'
       +DateToStr(Fiche.Q.FindField('E_DATECOMPTABLE').AsDateTime)+';'
       +FormatFloat('0000000000',Fiche.Q.FindField('E_NUMEROPIECE').AsInteger) ;
    LEcr.Add(St) ;
    Fiche.Q.Next ;
    end ;
  end else
    begin
    for i := 0 to Fiche.Fliste.NbSelected -1 do
      begin
      Fiche.FListe.GotoLeBookmark(i) ;
      {$IFDEF EAGLCLIENT}
      Fiche.Q.TQ.Seek(Fiche.FListe.Row-1);
      {$ENDIF}
      St:=Format_String(Fiche.Q.FindField('E_EXERCICE').AsString,3)+';'
       +Fiche.Q.FindField('E_JOURNAL').AsString+';'
       +Format_String(Fiche.Q.FindField('E_QUALIFPIECE').AsString,1)+';'
       +DateToStr(Fiche.Q.FindField('E_DATECOMPTABLE').AsDateTime)+';'
       +FormatFloat('0000000000',Fiche.Q.FindField('E_NUMEROPIECE').AsInteger) ;
      LEcr.Add(St) ;
      end ;
    end ;
End ;

Procedure TOF_CPPRORATATVA.ChargeTobGeneraux ;
Var Q : TQuery ;
Begin
TobGen := Tob.Create('_GENERAUX', Nil, -1) ;
//SG6 FQ 15155 11/01/05 Rajout du champ G_NATUREGENE
Q := OpenSql('SELECT G_GENERAL, G_NATUREGENE, G_TVA, G_REGIMETVA, G_TOTDEBE, G_TOTCREE, G_TOTDEBS, G_TOTCRES, '
           + 'G_TOTALDEBIT, G_TOTALCREDIT FROM GENERAUX WHERE G_NATUREGENE="CHA" OR G_NATUREGENE="IMO"'
           + ' OR G_NATUREGENE="PRO" OR G_GENERAL LIKE "445%"',True) ;
TobGen.LoadDetailDB('GENERAUX', '', '', Q, True) ;
Ferme(Q) ;
If TobGen.Detail.Count>0 then TobGen.Detail[0].AddChampSup('PRESENT',True);
End ;

Procedure TOF_CPPRORATATVA.ChargeTobDevise  ;
var Q : TQuery ;
BEGIN
TobDev := Tob.Create('_DEVISE',Nil,-1) ;
Q := OpenSql('SELECT D_DEVISE, D_DECIMALE FROM DEVISE', true) ;
TobDev.LoadDetailDB('DEVISE','','',Q,True,True) ;
Ferme(Q) ;
END ;

Procedure TOF_CPPRORATATVA.ChargeTobSection ;
Var Q:TQuery ;
BEGIN
TobSec := Tob.Create('_SECTION', Nil, -1) ;
Q := OpenSql('SELECT S_AXE, S_SECTION, S_TOTDEBE, S_TOTCREE, S_TOTDEBS, S_TOTCRES, S_TOTALDEBIT, S_TOTALCREDIT FROM SECTION',True) ;
TobSec.LoadDetailDB('SECTION', '', '', Q, False) ;
Ferme(Q) ;
If TobSec.Detail.Count>0 then TobSec.Detail[0].AddChampSup('PRESENT',True) ;
END ;

procedure TOF_CPPRORATATVA.ChargeTobProrata ;
var Q : TQuery ;
    st : String ;
    ch : String ;
BEGIN
TobProrata := Tob.create('_PRORATA', Nil, -1) ;
{JP 06/06/2003 : on autorise sur les exercices ouverts et maintenant provisoirement clôturés :
                 EX_ETATCPTA <> "CDE" au lieu de EX_ETATCPTA ="OUV"}
ch := RetourneClauseIn(GetControlText('E_JOURNAL'));
      {JP 10/06/2003 : Maintenant que l'on a un multiCombo, PA_JOURNAL doit être pris dans une liste de choix}
St := 'SELECT PA_CODE, PA_JOURNAL, PA_TAUX1, PA_TAUX2, PA_TAUX3, PA_DATEDEBUT, PA_DATEINTER FROM CPPRORATA ' +
      'LEFT JOIN EXERCICE ON PA_CODE=EX_EXERCICE WHERE EX_ETATCPTA <> "CDE" AND PA_CODE="'+GetControlText('E_EXERCICE')+'"';
// FQ 12913
if (ch <> '') then St := St + 'AND PA_JOURNAL IN ('+ ch +')';

if (Trim(ch) = '') or not ExisteSQL(St) then begin // S'il n'y a pas de taux paramètré pour le journal sélectionné, regarde s'il y en a un pour Tous les journaux
St := 'SELECT PA_CODE, PA_JOURNAL, PA_TAUX1, PA_TAUX2, PA_TAUX3, PA_DATEDEBUT, PA_DATEINTER FROM CPPRORATA '
     +'LEFT JOIN EXERCICE ON PA_CODE=EX_EXERCICE WHERE EX_ETATCPTA <> "CDE" AND PA_CODE="'+GetControlText('E_EXERCICE')+'"';  // *-*
// *+*     +'LEFT JOIN EXERCICE ON PA_CODE=EX_EXERCICE WHERE EX_ETATCPTA <> "CDE" AND PA_CODE="'+GetControlText('E_EXERCICE')+'" AND PA_TOUSJAL="X"';
end;

Q:=OpenSql(St, True) ;
TobProrata.LoadDetailDb('CPPRORATA', '', '', Q, True, True) ;
Ferme(Q) ;
END ;

Procedure TOF_CPPRORATATVA.FaitLeWhereTablesLibres (Var S : String ) ;
VAR i : Integer ;
begin
for i:=0 to 9 do
  begin
  if GetControlEnabled('G_TABLE'+IntToStr(i))= true then
    if GetControlText('G_TABLE'+IntToStr(i)) <>'' then S:=S+' AND ((G_TABLE'+IntToStr(i)+'="'+GetControlText('G_TABLE'+IntToStr(i))+'" AND (G_GENERAL LIKE "6%" OR G_GENERAL LIKE "7%")) OR ((G_GENERAL NOT LIKE "6%") OR (G_GENERAL NOT LIKE "7%")))' ;
  end ;
end ;

function TOF_CPPRORATATVA.ChargeTobEcr (i : integer; var OkOk : Boolean ; var Averto : Boolean) : boolean;
Var S, St, st2, StJal, StExo, StQual : String ;
    StDate : TDateTime ;
    StNum, j : Integer ;
    Q: Tquery ;
    TobGenDet : Tob ;
begin
//SG6 11/01/05 Gestion immo
Result := true;
TobEcr:=Tob.Create('_Ecr', Nil, -1) ;
S:=LEcr[i] ; StJal:='' ; StExo :='' ; StQual :='' ; StDate:=0 ; StNum :=0 ;
If S<>'' Then StExo:=ReadToKenST(S) ;
If S<>'' Then StJal:=ReadToKenST(S) ;
If S<>'' Then StQual:=ReadToKenST(S) ;
If S<>'' Then StDate:=StrToDate(ReadToKenST(S)) ;
If S<>'' Then StNum:=StrToInt(ReadToKenST(S)) ;
If (StJal='') Or (StExo='') or (StQual='') or (StDate=0) or (StNum=0) Then begin OkOk := False ; exit ; end ;
TobEcr.ClearDetail ;
St := 'SELECT E_JOURNAL, E_GENERAL, E_AUXILIAIRE, E_EXERCICE, E_DATECOMPTABLE, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE, '
    + 'E_NUMECHE, E_QUALIFPIECE, E_ETABLISSEMENT, E_TYPEMVT, E_REGIMETVA, E_TVA, E_DEBIT, E_CREDIT, E_DEVISE, '
    + 'E_DEBITDEV, E_CREDITDEV, E_ANA, E_NATUREPIECE, E_MODESAISIE, E_SEMAINE, '
    + 'E_QUALIFORIGINE,E_TAUXDEV, E_LIBELLE, E_REFINTERNE, E_CONTROLETVA, E_TVAENCAISSEMENT FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL '
    + 'WHERE E_JOURNAL="'+StJal+'" AND E_EXERCICE="'+StExo+'" AND E_QUALIFPIECE="'+StQual+'" AND '
    + 'E_DATECOMPTABLE="'+UsDateTime(StDate)+'" AND E_NUMEROPIECE='+IntToStr(StNum)+' AND (E_QUALIFORIGINE <> "TV2" OR E_QUALIFORIGINE IS NULL )'
    + 'AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="FF" OR E_NATUREPIECE="OD") ';
St2:='' ; FaitLeWhereTablesLibres (st2) ; St := St + St2 ;
if AvecEcrModifiees='Que' then St:=St+' AND (E_QUALIFORIGINE="TV1" OR E_QUALIFORIGINE="TVI")' ;
if AvecEcrModifiees='Sans' then st:=St+' AND ((E_QUALIFORIGINE<>"TV1" AND E_QUALIFORIGINE<>"TVI") Or (E_QUALIFORIGINE IS NULL))' ;
if AvecEcrModifiees='Indiff' then st:=St+'' ;
Q := OpenSql(St, True) ;
TobEcr.loadDetailDb('ECRITURE','','',Q,False,True) ;
Ferme(Q) ;
//SG6 11/01/05 Gestion des comptes immo FQ 15155
if TobEcr.Detail.Count > 0 then TobEcr.Detail[0].AddChampSup('IMO',true);
For j:=0 to TobEcr.Detail.Count-1 do
    BEGIN
    TobGenDet := TobGen.Findfirst(['G_GENERAL'],[TobEcr.Detail[j].GetValue('E_GENERAL')],False) ;

    if TobGenDet <> nil then
    begin
      TobGenDet.PutValue('PRESENT','X') ;
      //SG6 11/01/05 Gestion proratisation des immo FQ 15155
      if TobGenDet.GetString('G_NATUREGENE')='IMO' then
      begin
        tobEcr.Detail[j].PutValue('IMO','X');
        if Averto then
        begin
          if (HShowMessage('0;Attention;'+TraduireMemoire('Votre sélection contient des lignes d''immobilisation. Souhaitez-vous poursuivre le traitement ?')+';Q;YN;N;N;','','') <> mrYes) then 
          begin
            Result := False;
            Exit;
          end;
          Averto := False;
        end;
      end
      else
      begin
        tobEcr.Detail[j].PutValue('IMO','-');
      end;
    end;
  END ;
end ;

Function TOF_CPPRORATATVA.AvecEcrModifiees : String ;
Begin
Result:='Indiff' ;
If QuelProrata='First' then Result:='Sans' ; // sans les écritures
If QuelProrata='Last' then
  Begin
  If GetCheckBoxState('ECRMODIFIEES')=CbChecked then Result:='Que' else // que les ecr modifiées
    If GetCheckBoxState('ECRMODIFIEES')=CbUnchecked then Result:='Sans' ; // sans les écritures
  end ;
//If GetCheckBoxState('ECRMODIFIEES')=CbChecked then Result:='Que' ; // que les ecr modifiées
//If GetCheckBoxState('ECRMODIFIEES')=CbUnchecked then Result:='Sans' ; // sans les écritures
END ;

procedure TOF_CPPRORATATVA.AjouteLesAnas ;
Var i, j : integer ;
    TOBL, TobAna, TobSecDet : TOB ;
    SQL : String ;
    QQ   : TQuery ;
BEGIN
TobAna:=TOB.Create('',Nil,-1) ;
for i:=0 to TobEcr.Detail.Count-1 do
  BEGIN
  TOBL:=TOBEcr.Detail[i] ; if TOBL.GetValue('E_ANA')<>'X' then Continue ;
  SQL:='SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="'+TOBL.GetValue('E_JOURNAL')+'" '
      +'AND Y_EXERCICE="'+TOBL.GetValue('E_EXERCICE')+'" AND Y_DATECOMPTABLE="'+UsDateTime(TOBL.GetValue('E_DATECOMPTABLE'))+'" '
      +'AND Y_NUMEROPIECE='+IntToStr(TOBL.GetValue('E_NUMEROPIECE'))+' AND Y_NUMLIGNE='+IntToStr(TOBL.GetValue('E_NUMLIGNE'))+' '
      +'AND Y_QUALIFPIECE="N"' ;
  QQ:=OpenSQL(SQL,True) ;
  if Not QQ.EOF then TobAna.LoadDetailDB('ANALYTIQ','','',QQ,False,True) ;
  For j:=TobAna.Detail.Count-1 downto 0 do
    BEGIN
    TobSecDet := TobSec.Findfirst(['S_AXE','S_SECTION'],[TobAna.Detail[j].GetValue('Y_AXE'),
                                                         TobAna.Detail[j].GetValue('Y_SECTION')],False) ;
    if TobSecDet <> nil then TobSecDet.PutValue('PRESENT','X') ;
    TobAna.Detail[j].ChangeParent(TobEcr.Detail[i], -1) ;
    END ;
  Ferme(QQ) ;
  END ;
TobAna.Free ;
END ;

procedure TOF_CPPRORATATVA.InitTablesLibres ( TT : TTabSheet ) ;
Var LesLib : HTStringList ;
    i : Integer ;
    St,Titre,Ena : String ;
    Trouver : Boolean ;
    LL      : THLabel ;
    CC      : THEdit ;
BEGIN
if TT=Nil then Exit ; Trouver:=False ;
LesLib:=HTStringList.Create ;
GetLibelleTableLibre('G',LesLib) ;
for i:=0 to LesLib.Count-1 do
  BEGIN
  St:=LesLib.Strings[i] ; Titre:=ReadTokenSt(St) ; Ena:=St ;
  LL:=THLabel(TForm(TT.Owner).FindComponent('TG_TABLE'+IntToStr(i))) ;
  CC:=THEdit(TForm(TT.Owner).FindComponent('G_TABLE'+IntToStr(i))) ;
  if LL<>Nil then
    BEGIN
    LL.Caption:=Titre ; LL.Enabled:=(Ena='X') ;
    //if ((EstSerie(S3)) and (i>2)) then LL.Visible:=False ;
    if CC<>Nil then
      BEGIN
      CC.Enabled:=LL.Enabled ; if CC.Enabled then Trouver:=True ;
      //if ((EstSerie(S3)) and (i>2)) then CC.Visible:=False ;
      END ;
    END ;
  END ;
TT.TabVisible:=Trouver ;
LesLib.Clear ; LesLib.Free ;
END ;

Function TOF_CPPRORATATVA.RechercheCpteTVA (Ecr : Tob) : String ;
var TobGenDet : Tob ;
    Compte : String ;
    TVA, ModeTVA : String3 ;
BEGIN
TobGenDet := TobGen.FindFirst(['G_GENERAL'],[Ecr.GetValue('E_GENERAL')],False) ;
Compte := TobGenDet.GetValue('G_GENERAL') ; TVA := TobGenDet.GetValue('G_TVA') ;
ModeTVA := Ecr.GetValue('E_REGIMETVA') ;
if (Ecr.GetValue('E_TVAENCAISSEMENT') = 'X') then
  Result := TVA2ENCAIS(ModeTVA, TVA, True)
else
  Result := TVA2CPTE(ModeTVA, TVA, True) ;
END ;

Function TOF_CPPRORATATVA.RechercheTauxTVA (Ecr : TOB ) : Double ;
var TobGenDet : TOB ;
    Compte : String ;
    TVA, ModeTVA : String3 ;
BEGIN
TobGenDet := TobGen.FindFirst(['G_GENERAL'],[Ecr.GetValue('E_GENERAL')],False) ;
Compte := TobGenDet.GetValue('G_GENERAL') ;
TVA := TobGenDet.GetValue('G_TVA') ;
ModeTVA := Ecr.GetValue('E_REGIMETVA') ;
Result := TVA2TAUX(ModeTVA, TVA, True) ;
END ;

Function TOF_CPPRORATATVA.RecalculHorsTaxes(Ecr : Tob ; Tx1, TxTVA : Double; Z: String ; Dec: Integer) : Double ;
var HorsTaxes : Double ;
Begin
HorsTaxes := Arrondi((Ecr.GetValue('E_DEBIT'+Z)+ Ecr.GetValue('E_CREDIT'+Z)), dec2) ;
Result := Arrondi(HorsTaxes / (1+((1-(Tx1/100))*TxTVA)), dec) ;
End ;

Function TOF_CPPRORATATVA.CalculMontantTVA (Ecr : Tob ; Lib : String ; Tx1, Tx2, TxTVA : Double ; Dec : integer) : Double ;
var TobGenDet : TOB ;
    Compte,QualOrig : String ;
    TVA, ModeTVA : String3 ;
    HorsTaxes : Real ;
BEGIN
QualOrig := Ecr.GetValue('E_QUALIFORIGINE') ;
TobGenDet := TobGen.FindFirst(['G_GENERAL'],[Ecr.GetValue('E_GENERAL')],False) ;
Compte := TobGenDet.GetValue('G_GENERAL') ; TVA := TobGenDet.GetValue('G_TVA') ;
ModeTVA := Ecr.GetValue('E_REGIMETVA') ;

If (QualOrig = 'TV1') or (QualOrig = 'TVI') then
                         Horstaxes := RecalculHorsTaxes (Ecr, Tx1, TxTVA, Lib, Dec) else
                         HorsTaxes := Arrondi((Ecr.GetValue('E_DEBIT'+Lib)+ Ecr.GetValue('E_CREDIT'+Lib)), dec) ;
Result := HT2TVA(HorsTaxes, ModeTVA, False, TVA, '', True, dec) ;
END ;

Function TOF_CPPRORATATVA.RechercheLigneCpteTVA (Compte : String) : Integer ;
var j : Integer ;
BEGIN
Result :=-1 ;
For j:= 0 to TobEcr.Detail.Count-1 do if TobEcr.Detail[j].GetValue('E_GENERAL')=Compte then Result := j ;
END ;

Function TOF_CPPRORATATVA.RechercheTauxProrata (Ecr : TOB ; var QO : string ; Var T1, T2 :Double) : Double ;
Var TobProrataDet : TOB ;
    DateC, DateInter, DateFinale : TDateTime ;
    QA, Exo : String ;
BEGIN
Result := -1 ;
QA := Ecr.GetValue('E_QUALIFORIGINE') ;

// On cherche d'abord si la TOBPRORATA contient une tob fille qui aurait le même code journal que l'écriture
// Sinon on cherche la tob fille qui à le même code exercice
If Quelexo(Ecr.GetValue('E_DATECOMPTABLE'))=VH^.EnCours.Code then begin
  Datefinale := VH^.EnCours.Fin ;
  TobProrataDet := TobProrata.Findfirst(['PA_CODE','PA_JOURNAL'],[VH^.EnCours.Code,Ecr.GetValue('E_JOURNAL')],False);
  if (TobProrataDet = nil) then TobProrataDet := TobProrata.Findfirst(['PA_CODE'],[VH^.EnCours.Code],False);
  end
else begin
  Datefinale := VH^.Suivant.Fin ;
  TobProrataDet := TobProrata.Findfirst(['PA_CODE','PA_JOURNAL'],[VH^.Suivant.Code,Ecr.GetValue('E_JOURNAL')],False);
  if (TobProrataDet = nil) then TobProrataDet := TobProrata.Findfirst(['PA_CODE'],[VH^.Suivant.Code],False);
end;

//Si Oeuvres laïques voir pour journal à ---
// Si pas trouvé : TobProrataDet = nil
if (TobProrataDet = nil) then exit;
DateInter := TobProrataDet.GetValue('PA_DATEINTER') ; DateC := Ecr.GetValue('E_DATECOMPTABLE') ;
Exo := Quelexo(Ecr.GetValue('E_DATECOMPTABLE')) ;
if QA='TV1' then T1 := TobProrataDet.GetValue('PA_TAUX1') ;

if QA='TVI' then T1 := TobProrataDet.GetValue('PA_TAUX3') ;
T2 := TobProrataDet.GetValue('PA_TAUX2') ;

if (TobProrataDet.GetValue('PA_TAUX2') <> 0) then  // Si le taux DEFINITIF (PA_TAUX2) est renseigné
  if (DateC > DateInter) or (DateInter=datefinale) then
    Begin Result := TobProrataDet.GetValue('PA_TAUX2') ; Qo := 'TV2' ; exit ; end ;

if (TobProrataDet.GetValue('PA_TAUX3') <> 0) then  // Si le taux INTERMEDIAIRE (PA_TAUX3) est renseigné
  if QA='' then
    Begin Result := TobProrataDet.GetValue('PA_TAUX3') ; Qo := 'TVI' ; exit ; end else
    Begin Result := -1 ; exit ; end ;

if (TobProrataDet.GetValue('PA_TAUX1')<>0) then // Si le premier Taux est renseigné
  BEGIN
  if QA='TV1' then begin Result := -1 ; exit ; end else
                   begin Result:=TobProrataDet.GetValue('PA_TAUX1') ; QO:='TV1' ; end ;
  END ;
END ;


Function TOF_CPPRORATATVA.RechDecDev (Ecr : TOB) : Integer ;
Var TobDecDev : TOB ;
BEGIN
Result := 4 ;
TobDecDev := TobDev.Findfirst(['D_DEVISE'], [Ecr.GetValue('E_DEVISE')], False) ;
If TobDecDev <> Nil then Result := TobDecDev.getValue('D_DECIMALE') ;
END ;

Procedure TOF_CPPRORATATVA.MajCompte(Zone : String ; E : Double ; Exo : String ; TobHT, TobTVA : TOB) ;
var TotalZone, TotalZoneTVA, ZoneTVA : String ;
    TobGenHT, TobGenTVA : TOB ;
BEGIN
if Zone='E_DEBIT' then
  begin
  TotalZone:='G_TOTALDEBIT' ; TotalZoneTVA:='G_TOTALCREDIT' ;
  if Exo=VH^.EnCours.Code then begin Zone:='G_TOTDEBE' ; ZoneTVA := 'G_TOTCREE' ; end
                          else begin Zone:='G_TOTDEBS' ; ZoneTVA := 'G_TOTCRES' ; end ;
  end else
  BEGIN
  TotalZone:='G_TOTALCREDIT' ; TotalZoneTVA:='G_TOTALDEBIT' ;
  if Exo=VH^.EnCours.Code then begin Zone:='G_TOTCREE' ; ZoneTVA := 'G_TOTDEBE'; end
                          else begin Zone:='G_TOTCRES' ; ZoneTVA := 'G_TOTDEBS' ; end ;
  end ;

//MàJ ligne Hors Taxes
TobGenHT := TobGen.FindFirst(['G_GENERAL'],[TobHT.GetValue('E_GENERAL')],False) ;
TobGenHT.PutValue(Zone, Arrondi(TobGenHT.GetValue(Zone)+E, Dec2)) ;
TobGenHT.PutValue(TotalZone, Arrondi(TobGenHT.GetValue(TotalZone)+E, Dec2)) ;
//MàJ ligne de TVA
If TobHt.GetValue('E_QUALIFORIGINE') = 'TV2' then
  begin
  TobGenTVA := TobGen.FindFirst(['G_GENERAL'],[TobTVA.GetValue('E_GENERAL')],False) ;
  TobGenTVA.PutValue(ZoneTVA, Arrondi(TobGenTVA.GetValue(ZoneTVA)-E, Dec2)) ;
  TobGenTVA.PutValue(TotalZoneTVA, Arrondi(TobGenTVA.GetValue(TotalZoneTVA)-E, Dec2)) ;
  end else
  begin
  TobGenTVA := TobGen.FindFirst(['G_GENERAL'],[TobTVA.GetValue('E_GENERAL')],False) ;
  TobGenTVA.PutValue(Zone, Arrondi(TobGenTVA.GetValue(Zone)-E, Dec2)) ;
  TobGenTVA.PutValue(TotalZone, Arrondi(TobGenTVA.GetValue(TotalZone)-E, Dec2)) ;
  end ;
END ;

procedure TOF_CPPRORATATVA.MajAnalytiq (var TobHT : TOB ; DecDev : Integer ; Zone, Exo : String) ;
Var ZoneA : String ;
    AR, ARD : Double ; // Montant à répartir en Frs, devises
    M, Morig, MD, Ecart : Double ; // Montant de la ligne en F,D et Ecart entre montant d'origine et recalculé
    CM, CMD : Double ; // Cumul montant F,D
    TobAnaDet, TobSecDet : TOB ;
    i : integer ;
BEGIN

AR :=TobHT.GetValue('E_'+Zone) ; ARD:=TobHT.GetValue('E_'+Zone+'DEV') ;

CM := 0 ; CMD := 0 ; // RàZ des cumuls
for i :=0 to TobHT.Detail.Count-1 do
  Begin
  TobAnaDet := TobHT.Detail[i] ; MOrig := TobAnaDet.GetValue('Y_'+Zone) ;

  // Remet à 0 les compteurs si on change d'axe (dans le cas d'une ventilation analytique sur plusieurs axe) VL - FQ 12646
  if (i > 0) then
    if (TobAnaDet.GetValue('Y_AXE') <> TobHT.Detail[i-1].GetValue('Y_AXE')) then begin
      CM := 0;
      CMD := 0;
    end;

  M := Arrondi(TobAnaDet.GetValue('Y_POURCENTAGE')*AR/100, Dec2) ; CM := CM+M ;
  MD := Arrondi(TobAnaDet.GetValue('Y_POURCENTAGE')*ARD/100, DecDev) ; CMD := CMD + MD ;

  // si dernière ligne et total different d'un des cumuls (Frs, Devises, Euros)
  If ((i=TobHT.Detail.Count-1) and ((CM <> AR) or (CMD <> ARD))) then
    begin
    M := Arrondi(M + AR - CM, Dec2) ;
    MD := Arrondi(MD + ARD - CMD, DecDev) ;
    end ;

  {JP 01/06/05 : FQ 15960 : On réinitialise les montants à zéro, car TobAnaDet est issu d'une duplication,
                 donc les Debits ou les Credits sont renseignés donc si Zone est différente des colonnes
                 remplies ....
                 Ex : Y_DEBIT = 254.32 MD = 15 et Zone = CREDIT, on se retrouve avec Y_DEBIT = 254.32 et Y_CREDIT = 15}
  TobAnaDet.SetDouble('Y_DEBIT' , 0.0);
  TobAnaDet.SetDouble('Y_CREDIT', 0.0);
  TobAnaDet.SetDouble('Y_DEBITDEV' , 0.0);
  TobAnaDet.SetDouble('Y_CREDITDEV', 0.0);

  TobAnaDet.PutValue('Y_'+Zone, M) ;
  TobAnaDet.PutValue('Y_'+Zone+'DEV', MD) ;

  TobAnaDet.PutValue('Y_TOTALECRITURE', AR) ;
  TobAnaDet.PutValue('Y_TOTALDEVISE', ARD) ;

  TobSecdet := TobSec.FindFirst(['S_AXE','S_SECTION'],
                                [TobAnaDet.GetValue('Y_AXE'),TobAnaDet.GetValue('Y_SECTION')],False) ;

  if Exo = VH^.EnCours.Code then ZoneA:='S_TOT'+Copy(Zone, 1, 3)+'E' else ZoneA:='S_TOT'+Copy(Zone, 1, 3)+'S' ;
  Ecart := Arrondi(M-MOrig, Dec2) ; TobSecDet.PutValue(ZoneA, Arrondi(TobSecDet.GetValue(ZoneA)+Ecart, Dec2)) ;
  End ;
END ;


Procedure TOF_CPPRORATATVA.ModifLignes (var TobTTC, TobHT, TobTVA : TOB ; TxP, TxTVA, MtTVA, MtTVAD : Double ; QO : String) ;
var MtTVAOrig, MtTVADOrig, TVA, TVAD : Double ;
    Ecart, EcartD : Double ;
    Zone, QO2 : String ;
    DecDev, NumTVA : Integer ;
BEGIN

// Mémorise le montant HT initial
MemoriseHTInitial(TobHT);

DecDev :=RechDecDev(TobTVA) ;
If abs(TobTVA.GetValue('E_DEBIT'))>0 then Zone := 'DEBIT' else Zone := 'CREDIT' ;

MtTVAorig := TobTVA.GetValue('E_'+Zone); MtTVADOrig := TobTVA.GetValue('E_'+Zone+'DEV');
TVA := Arrondi(MtTVA*TxP/100, Dec2) ; TVAD := Arrondi(MtTVAD*TxP/100, DecDev) ;
Ecart := MtTVA-TVA ; EcartD := MtTVAD - TVAD ;

TobTVA.PutValue('E_'+Zone,Arrondi(MtTVAOrig-Ecart, Dec2)); TOBHT.PutValue('E_'+Zone,Arrondi(TobHT.GetValue('E_'+Zone)+Ecart, Dec2)) ;
TobTVA.PutValue('E_'+Zone+'DEV',Arrondi(MtTVADOrig-EcartD, DecDev)); TOBHT.PutValue('E_'+Zone+'DEV',Arrondi(TobHT.GetValue('E_'+Zone+'DEV')+EcartD, DecDev)) ;

TobTTC.PutValue('E_QUALIFORIGINE', QO) ; TobHt.PutValue('E_QUALIFORIGINE', QO) ; TobTVA.PutValue('E_QUALIFORIGINE', QO) ; // maj de l'ecriture avec e_qualiforigine

if (QO = 'TV1') then QO2:='PR1'
else
  if (QO = 'TVI') then QO2:='PR2'
  else
    if (QO = 'TV2') then QO2:='PR3'
    else
      QO2:='MTV';

TobTTC.PutValue('E_CONTROLETVA', QO2) ; TobHt.PutValue('E_CONTROLETVA', QO2) ; TobTVA.PutValue('E_CONTROLETVA', QO2) ; {18.07.07&07.11.07 FQ20654 YMO}

// VL - FQ 10486
NumTVA := Tva2NumBase(TobHT.GetValue('E_TVA'));
if NumTVA > 0 then begin
  TobTTC.PutValue('E_ECHEENC'+IntToStr(NumTVA), TOBHT.GetValue('E_'+Zone));
end;

MajCompte('E_'+Zone, Ecart, QuelExo(TobTVA.GetValue('E_DATECOMPTABLE')), TobHT, TobTVA) ; // Mise à jour comptes comptables
if TobHT.GetValue('E_ANA') = 'X' Then
begin
MajAnalytiq(TobHT, DecDev, Zone, QuelExo(TobHT.GetValue('E_DATECOMPTABLE'))) ; // mise à jour des lignes et sections analytiques
MajAnalytiq(TobTVA, DecDev, Zone, QuelExo(TobTVA.GetValue('E_DATECOMPTABLE'))) ; {FQ21704 19.11.2007  YMO Maj de la ligne ana su compte TVA}
end;
TobTTC.UpdateDb ; TobTVA.UpdateDB ; TobHT.UpdateDB ;
END ;

Procedure TOF_CPPRORATATVA.NouvelLignes (var TobTTC, TobHT, TobTVA : TOB ; TxP, TxTVA, MtTVA, MtTVAD : Double ; QO : String) ;
var Ecart, EcartD, TxOrig : Double ; // Montant final de la nouvelle écriture
    ZoneTVA, ZoneHT : String ;
    DecDev, NumTVA : Integer ;
    TobProrataDet : TOB ;
BEGIN
DecDev :=RechDecDev(TobTVA) ;
If TobTVA.GetValue('E_DEBIT')>0 then
                                begin ZoneTVA := 'DEBIT' ; ZoneHT := 'CREDIT' ; end
                           else begin ZoneTVA := 'CREDIT' ; ZoneHT := 'DEBIT' ; end ;

TobTVA.PutValue('E_DEBIT', 0); TobTVA.PutValue('E_CREDIT', 0) ;
TobTVA.PutValue('E_DEBITDEV', 0) ; TobTVA.PutValue('E_DEBITDEV', 0) ;
TobHT.PutValue('E_DEBIT', 0) ; TobHT.PutValue('E_CREDIT', 0) ;
TobHT.PutValue('E_DEBITDEV', 0) ; TobHT.PutValue('E_DEBITDEV', 0) ;

TobProrataDet := TobProrata.Findfirst(['PA_CODE','PA_JOURNAL'],[QuelExo(TobHT.GetValue('E_DATECOMPTABLE')),TobHT.GetValue('E_JOURNAL')],False) ;
if (TobProrataDet = nil) then
  TobProrataDet := TobProrata.Findfirst(['PA_CODE'],[QuelExo(TobHT.GetValue('E_DATECOMPTABLE'))],False) ;

If TobHT.GetValue('E_QUALIFORIGINE')='TV1' then TxOrig := TobProrataDet.GetValue('PA_TAUX1') else TxOrig := TobProrataDet.GetValue('PA_TAUX3') ;

Ecart := abs(Arrondi(MtTva * TxOrig / 100, Dec2)-(Arrondi(MtTVA * TxP / 100, Dec2))) ;
EcartD := abs(Arrondi(MtTvaD * TxOrig / 100, DecDev)-(Arrondi(MtTVAD * TxP / 100, DecDev))) ;

TobTVA.PutValue('E_'+ZoneTVA, Arrondi(Ecart, Dec2));
ToBHT.PutValue('E_'+ZoneHT, Arrondi(Ecart, Dec2)) ;

TobTVA.PutValue('E_'+ZoneTVA+'DEV',Arrondi(EcartD, DecDev));
ToBHT.PutValue('E_'+ZoneHT+'DEV',Arrondi(EcartD, DecDev)) ;

TobHt.PutValue('E_QUALIFORIGINE', QO) ;
TobTVA.PutValue('E_QUALIFORIGINE', QO) ; // maj de l'ecriture avec e_qualiforigine
{18.07.07 FQ20654 YMO}
TobHt.PutValue('E_CONTROLETVA', QO) ;
TobTVA.PutValue('E_CONTROLETVA', QO) ;

// VL - FQ 10486
NumTVA := Tva2NumBase(TobHT.GetValue('E_TVA'));
if NumTVA > 0 then begin
  TobTTC.PutValue('E_ECHEENC'+IntToStr(NumTVA), TOBHT.GetValue('E_'+ZoneHT));
end;

MajCompte('E_'+ZoneHT, Ecart, QuelExo(TobTVA.GetValue('E_DATECOMPTABLE')), TobHT, TobTVA) ; // Mise à jour comptes comptables
if TobHT.GetValue('E_ANA') = 'X' Then
begin
MajAnalytiq(TobHT, DecDev, ZoneHT, QuelExo(TobHT.GetValue('E_DATECOMPTABLE'))) ; // mise à jour des lignes et sections analytiques
MajAnalytiq(TobTVA, DecDev, ZoneHT, QuelExo(TobTVA.GetValue('E_DATECOMPTABLE'))) ; {FQ21704 19.11.2007  YMO Maj de la ligne ana su compte TVA}
end;

TobTVA.InsertDB(Nil) ;
TobHT.InsertDB(Nil) ;
END ;

Procedure TOF_CPPRORATATVA.ModifEcrAna (Var TobE : Tob ; Zone1, Zone2  : String ; NumPiece, Numligne : Integer; Dat : TDateTime) ;
var i : integer ;
begin
For i := 0 to TobE.Detail.Count-1 do
  Begin
  TobE.Detail[i].PutValue('Y_NUMEROPIECE', NumPiece) ;
  TobE.Detail[i].PutValue('Y_NUMLIGNE', NumLigne) ;
  TobE.Detail[i].PutValue('Y_TOTALECRITURE', arrondi(TobE.Detail[i].GetValue('Y_DEBIT')+TobE.Detail[i].GetValue('Y_CREDIT'), Dec2)) ;
  TobE.Detail[i].PutValue('Y_TOTALDEVISE', arrondi(TobE.Detail[i].GetValue('Y_DEBITDEV')+TobE.Detail[i].GetValue('Y_CREDITDEV'), Dec2)) ;
  TobE.Detail[i].PutValue('Y_'+Zone1, TobE.Detail[i].GetValue('Y_'+Zone2)) ;
  TobE.Detail[i].PutValue('Y_'+Zone2, 0) ;
  TobE.Detail[i].PutValue('Y_'+Zone1+'DEV', TobE.Detail[i].GetValue('Y_'+Zone2+'DEV')) ;
  TobE.Detail[i].PutValue('Y_'+Zone2+'DEV', 0) ;
  TobE.Detail[i].PutValue('Y_NATUREPIECE','OD') ;
  TobE.Detail[i].PutValue('Y_DATECOMPTABLE', Dat) ;
  TobE.Detail[i].PutValue('Y_SEMAINE', NumSemaine(TobE.Detail[i].GetValue('Y_DATECOMPTABLE'))) ;
  TobE.Detail[i].PutValue('Y_EXPORTE','---') ;
  TobE.Detail[i].PutValue('Y_CREERPAR','SAI') ;
  End ;
End ;

Function TOF_CPPRORATATVA.RechercheDate(TobE : TOB) : TDateTime ;
BEGIN
Result:=(StrToDate(Dategen.text)) ;
END ;

Procedure TOF_CPPRORATATVA.MajZones (Var TobE : TOB ; DD : TdateTime ; NumPiece, Numlig : Integer ; pszAuxi : String  ) ;
var
  Q : TQuery;
  szAuxi : String;
Begin
  Q := OpenSql('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+pszAuxi+'"',True);
  szAuxi := Q.Fields[0].AsString;
  Ferme(Q);
TobE.PutValue('E_REFINTERNE', 'Régul TVA n° '+IntToStr(TobE.GetValue('E_NUMEROPIECE'))+' du '+DateToStr(TobE.GetValue('E_DATECOMPTABLE'))+'') ;
TobE.PutValue('E_LIBELLE', Copy('Prorata TVA : '+szAuxi,0,35)) ;
TobE.PutValue('E_DATECOMPTABLE', DD) ;
TobE.PutValue('E_NUMEROPIECE', NumPiece ) ;
TobE.PutValue('E_NUMLIGNE',Numlig) ;
TobE.PutValue('E_SEMAINE', NumSemaine(TobE.GetValue('E_DATECOMPTABLE'))) ;
TobE.PutValue('E_NATUREPIECE','OD') ;
TobE.PutValue('E_DEVISE',V_PGI.DevisePivot) ;
TobE.PutValue('E_ETATLETTRAGE','RI') ;
TobE.PutValue('E_ECHE', '-') ;
TobE.PutValue('E_NUMECHE', 0) ;
TobE.PutValue('E_COUVERTURE', 0) ;
TobE.PutValue('E_COUVERTUREDEV', 0) ;
TobE.PutValue('E_ECRANOUVEAU','N') ;
TobE.PutValue('E_EXPORTE','---') ;
TobE.PutValue('E_CREERPAR','SAI') ;
TobE.PutValue('E_ENCAISSEMENT','RIE') ;
TobE.PutValue('E_CONTREPARTIEGEN','RIE') ;
End ;

function TOF_CPPRORATATVA.DupliqueLignes (var TobHT, TobTVA : TOB; pszAuxi : String) : Double ;
var NumPiece : Integer ;
    Qo : String ;
    Tx1, Tx2 : Double ;
    DD: TDateTime ;
BEGIN
NumPiece := GetNewNumJal(TobHT.GetValue('E_JOURNAL'),True,Date) ;
DD := RechercheDate (TobHt) ;
Tx1 := 0 ; Tx2 := 0 ; Qo := '' ; Result := RechercheTauxProrata(TobHT, QO, Tx1, Tx2) ;
if (Result = -1) then begin EcrNonTraitee:=True ; Exit ;end ;
if (Tx1>Tx2) and (TobHt.GetValue('E_DEBIT') > 0) then
    Begin
    TobTVA.PutValue('E_CREDIT', TobTVA.GetValue('E_DEBIT')) ; TobTVA.PutValue('E_DEBIT', 0) ;
    TobTVA.PutValue('E_CREDITDEV', TobTVA.GetValue('E_DEBITDEV')) ; TobTVA.PutValue('E_DEBITDEV', 0) ;
    if TobHT.GetValue('E_ANA') = 'X' then ModifEcrAna (TobHT, 'CREDIT', 'DEBIT', NumPiece, 1, DD) ;
    end else
    begin
    TobHT.PutValue('E_CREDIT', TobHT.GetValue('E_DEBIT')) ; TobHT.PutValue('E_DEBIT', 0) ;
    TobHT.PutValue('E_CREDITDEV', TobHT.GetValue('E_DEBITDEV')) ; TobHT.PutValue('E_DEBITDEV', 0) ;
    if TobHT.GetValue('E_ANA') = 'X' then ModifEcrAna (TobHT, 'DEBIT', 'CREDIT', NumPiece, 1, DD) ;
    end ;

MajZones(TobHT, DD, Numpiece, 1, pszAuxi) ;
MajZones(TobTVA, DD, Numpiece, 2, pszAuxi) ;
END ;

// Contrôles et vérifications
function TOF_CPPRORATATVA.VerifieLaPiece : boolean ;
var LigneTTC, i : integer ;
    OkHT, OkTVA : Boolean ;
//    RegimeTva, Etva: String ;
begin
Result :=True ;
LigneTTC := 0 ; OkHT := False ; OkTVA := False ;
//RegimeTva:=TobEcr.Detail[0].GetValue('E_REGIMETVA') ;
//ETva:=TobEcr.Detail[0].GetValue('E_TVA') ;
for i:=0 to TobEcr.Detail.count-1 do begin // Contrôles et vérifications
    If TobEcr.Detail[i].GetValue('E_MODESAISIE')<> '-' then begin
      MsgBox.Execute(11, TobEcr.Detail[i].GetString('E_NUMEROPIECE'), ''); // L'écriture n'est pas en mode pièce.
      Result := False ;
      exit ;
    end ;
    If TobEcr.Detail[i].GetValue('E_QUALIFORIGINE')= 'TV2' then begin
      MsgBox.Execute(12, TobEcr.Detail[i].GetString('E_NUMEROPIECE'), ''); // L'écriture a déjà été proratisée avec le taux final.
      Result := False ;
      exit ;
    end ;
    If (TobEcr.Detail[i].GetValue('E_NATUREPIECE')<>'AF') and (TobEcr.Detail[i].GetValue('E_NATUREPIECE')<>'FF') then begin
      MsgBox.Execute(13, TobEcr.Detail[i].GetString('E_NUMEROPIECE'), ''); // La piece n'est pas de nature Avoir ou Facture
      Result := False ;
      exit ;
    end ;
  //  if (TobEcr.detail[i].GetValue('E_REGIMETVA')<> RegimeTva) or
  //     (TobEcr.detail[i].GetValue('E_TVA')<> ETva) then begin result:=False ; exit ; end ;// +sieurs régimes de TVA differents
  
  if (TobEcr.Detail[i].GetValue('E_TYPEMVT')='TTC')
  {FQ22044  17.12.2007  YMO Prise en commpte des pièce multi-échéances}
  and (TobEcr.Detail[i].GetValue('E_NUMECHE')<=1) then LigneTTC := LigneTTC + 1 ;
  //SG6 11/01/05 FQ 15155 on considere que les comptes imo sont des comptes HT (pour un imo Y_TYPEMVT = DIV)
  if (TobEcr.Detail[i].GetValue('E_TYPEMVT')='HT') or (TobEcr.Detail[i].GetValue('IMO')='X') then OkHt := True ;

  if TobEcr.Detail[i].GetValue('E_TYPEMVT')='TVA' then OkTVA := True ;
    // Pas de ligne Ht ou TVA ou ligne de TTC <> 1 donc pas de traitement possible
    if (LigneTTC > 1) or (((LigneTTC=0) or (not OkHT) or (not OkTVA)) and (i=TobEcr.Detail.Count-1)) then begin
      if (LigneTTC >1) then MsgBox.Execute(14, TobEcr.Detail[i].GetString('E_NUMEROPIECE'), '') else // La piece %% a plusieurs lignes de TTC.
      if (LigneTTC =0) then MsgBox.Execute(15, TobEcr.Detail[i].GetString('E_NUMEROPIECE'), '') else // La piece %% n'a pas de ligne de TTC.
      if (not OkHT) then MsgBox.Execute(16, TobEcr.Detail[i].GetString('E_NUMEROPIECE'), '') else // La piece %% n'a pas de ligne de HT.
      if (not OkTVA) then MsgBox.Execute(17, TobEcr.Detail[i].GetString('E_NUMEROPIECE'), ''); // La piece %% n'a pas de ligne de TVA.
      Result := False;
      exit;
    end;
  end;
end ;

Procedure TOF_CPPRORATATVA.Traitement ;
var i, LigneTVA, LigneTC : integer ;
    EcritureDejaModifiee : Boolean ;
    CpteTVA, QualifOrigine, szAuxi {, RegimeTva, Etva} : String ;
    TobHTini, TobTVAini, TobHT2, TobTVA2, TobTTC : TOB ;
    TauxTVA,TauxProrata, MontantTVA, MontantTVADEV, Tx1, Tx2 : Double ;
BEGIN
if (not VerifieLaPiece) then exit ;

for i:=0 to TobEcr.Detail.Count-1 do // Traitement de la pièce
  BEGIN
  //SG6 11/01/05 FQ 15155 on considere que les comptes imo sont des comptes HT (pour un imo Y_TYPEMVT = DIV)
  if (TobEcr.Detail[i].GetValue('E_TYPEMVT')='HT') or (TobEcr.Detail[i].GetValue('IMO')='X') then
    BEGIN
    CpteTVA := RechercheCpteTVA (TobEcr.Detail[i]) ; LigneTva := RechercheLigneCpteTVA(CpteTVA) ;
    if LigneTVA = -1 then begin EcrNonTraitee:=True ; continue ; end ; // si la ligne de TVA n'existe pas
    TobHTini := TobEcr.Detail[i] ; TobTVAini := TobEcr.Detail[LigneTVA] ;

    {LigneTc:=0 ; } LigneTc := RechercheLigneTTC ; TobTTc:=TobEcr.Detail[LigneTC] ;

    TauxTVA := RechercheTauxTVA (TobHTini) ; QualifOrigine :='' ;
    TauxProrata := RechercheTauxProrata (TobHTini, Qualiforigine, tx1, tx2) ;
    if TauxProrata = -1 then begin EcrNonTraitee:=True ; exit ; end ;
    MontantTVA := CalculMontantTva (TobHTini,'', Tx1, Tx2, TauxTVA, Dec2) ;
    MontantTVADev := CalculMontantTva (TobHTini,'DEV', Tx1, Tx2, TauxTVA, RechDecDev(TobHTini)) ;

    EcritureDejaModifiee :=(TobEcr.Detail[i].GetValue('E_QUALIFORIGINE')='TV1') or
       (TobEcr.Detail[i].GetValue('E_QUALIFORIGINE')='TVI') ;

    if EcritureDejaModifiee then
      begin
      if TobHt2 = nil then TobHT2 := Tob.Create('HT', nil , -1) ;
      TobHT2.Dupliquer(TobHTini, True, True) ;
      if TobTVA2 = nil then TobTVA2 := Tob.Create('TVA', nil , -1) ;
      TobTVA2.Dupliquer(TobTVAini, True, True) ;
      TobTVAini.PutValue('E_QUALIFORIGINE','TV2') ;TobHTini.PutValue('E_QUALIFORIGINE','TV2') ; TobTTC.PutValue('E_QUALIFORIGINE','TV2') ;
      TobTVAini.UpdateDb; TobHTIni.UpdateDB ;
      if (DupliqueLignes(TobHT2, TobTVA2,szAuxi) = -1) then begin exit ; end ;

      NouvelLignes(TobTTC, TobHT2, TobTVA2, TauxProrata, TauxTVA, MontantTVA, MontantTVADev, QualifOrigine) ;// création ligne de régularisation HT <---> TVA
      TobTVA2.Free ; TobTVA2 := Nil ; TobHT2.Free ; TobHT2 := Nil ;
      end else ModifLignes (TobTTC, TobHTini, TobTVAini, TauxProrata, TauxTVA, MontantTVA, MontantTVADev, QualifOrigine) ; // Modification de l'écriture proratisée
    END ;
  END ;
END ;

Function TOF_CPPRORATATVA.RechercheLignettC : Integer ;
Var j : Integer ;
Begin
  Result := 0;
  For j:= 0 to TobEcr.Detail.Count-1 do
    if TobEcr.Detail[j].GetValue('E_TYPEMVT')='TTC' then Result := j ;
end ;

Procedure TOF_CPPRORATATVA.MajTobs ;
var i : integer ;
begin
For i := 0 to TobGen.Detail.Count-1 do
  Begin
  If TobGen.Detail[i].GetValue('PRESENT')='X' then TobGen.Detail[i].UpdateDb ;
  End ;
For i := 0 to TobSec.Detail.Count-1 do
  Begin
  If TobSec.Detail[i].GetValue('PRESENT')='X' then TobSec.Detail[i].UpdateDb ;
  End ;
end ;

Procedure TOF_CPPRORATATVA.BOuvrirOnClick (Sender : TObject) ;
var i : Integer ;
    OkOk : Boolean ;
    bMessage : Boolean ;
    Q: TQuery ;
    TobProrataDet : TOB ;
    bAverto : boolean;
Begin
If (MsgBox.Execute(0, '', '') = MrNo) then exit; // Confirmez-vous le traitement ?

{$IFDEF EAGLCLIENT}
if Fiche.bSelectAll.Down then
  if not Fiche.FetchLesTous then
  begin
    Fiche.bSelectAllClick(nil);
    Fiche.bSelectAll.Down := False;
    Exit;
  end;
{$ENDIF}

Q := OpenSql('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exercice.Value+'"', True) ;
If ((StrToDate(DateGen.Text) < Q.Fields[0].AsDateTime) or (StrToDate(DateGen.Text) > Q.Fields[1].AsDateTime)) then begin
  MsgBox.Execute(9, '', ''); // La date de génération des écritures n''est pas inclue dans l''exercice.
  Dategen.SetFocus ; Ferme(Q) ; exit ;
end ;
Ferme(Q) ;

If QuelProrata='Last' then if (not VerifieTauxDefinitifSaisi) then exit ;

EnableControls(Ecran, False ) ;

// Vérification qu'une ligne au moins a été selectionnée
if (ListeEcr.NbSelected=0) and (Not ListeEcr.AllSelected) then begin
  MsgBox.Execute(1, '', ''); // Aucune ligne sélectionnée. Traitement impossible.
  EnableControls(Ecran, True) ; Exit ;
end;

Chargement ; If Lecr.Count<=0 Then begin EnableControls(Ecran, True) ;Exit ;end ;

{$IFDEF EAGLCLIENT}
Fiche.Q.TQ.Close;
{$ELSE}
Fiche.Q.Close ;
{$ENDIF}

// Récup comptes généraux + devises + sections + prorata de tva
ChargeTobProrata ; ChargeTobGeneraux ; ChargeTobDevise ; ChargeTobSection ;

// Vérifie si les taux de prorata sont renseignés
if TobProrata.Detail.Count = 0 then begin
  MsgBox.Execute(2, '', ''); // Veuillez paramétrer les taux de prorata de TVA.
  LibereTobs; EnableControls(Ecran, True); Exit;
end;

// S'assure que le taux définitif est renseigné
if (QuelProrata='Last') then begin
  bMessage := False;
  for i := 0 to TobProrata.Detail.Count-1 do if (TobProrata.Detail[0].GetValue('PA_TAUX2') = 0) then bMessage := True;

  if bMessage then begin
    MsgBox.Execute(3, '', ''); // Veuillez paramétrer le taux de prorata de TVA définitif pour l''année sélectionnée.
    LibereTobs; EnableControls(Ecran, True); Exit;
  end;
end;

// Vérifie si le taux intermédiaire et définitif sont différents (s'ils sont identiques,on obtient des écritures à 0)
TobProrataDet := TobProrata.Findfirst(['PA_CODE','PA_JOURNAL'],[GetControlText('E_EXERCICE'),GetControlText('E_JOURNAL')],False);
if (TobProrataDet = nil) then
  TobProrataDet := TobProrata.Findfirst(['PA_CODE'],[GetControlText('E_EXERCICE')],False);
if (TobProrataDet = nil) then begin LibereTobs; EnableControls(Ecran, True); Exit; end;

{  // FQ 12252 Point 1
if (TobProrataDet.GetValue('PA_TAUX2') = TobProrataDet.GetValue('PA_TAUX3')) then begin
  Msgbox.Execute(4, '', ''); // Le taux intermédiaire et définitif sont identiques pour l''année sélectionnée.#10#13Veuillez revoir le paramétrage des taux de prorata de TVA.
  LibereTobs; EnableControls(Ecran, True); Exit;
end;}

bAverto := True;
i := 0;
while (i <= LEcr.count-1) and (bAverto = True) do
begin
    if not ChargeTobEcr(i,OkOk,bAverto) then
    begin
      Liberetobs ;
      {$IFDEF EAGLCLIENT}
       Fiche.Q.TQ.Open;
      {$ELSE}
       Fiche.Q.Open ;
      {$ENDIF}
            TToolBarButton97(GetControl('BCHERCHE')).Click ;
      EnableControls(Ecran, True) ;
      Exit;
    end;
    Inc(i);
end;

Try
  BeginTrans ;
  EcrNonTraitee:=False ;
  For i := 0 to LEcr.Count-1 do
    BEGIN
    OkOk:= True ;
    ChargeTobEcr(i,OkOk,bAverto);
    if not OkOk then continue ;
    AjouteLesAnas ;
    Traitement ;
    END ;
  MajTobs ;
  CommitTrans ;
  If EcrNonTraitee then
  begin
    MsgBox.Execute(5, '', '')   // ATTENTION : Certaines pièces n''ont pas été proratisées. Lancez le contrôle des factures.
  end
  else
  begin
    if not bAverto then PGIBox('Attention le traitement de proratisation des lignes d''immobilisation n''est pas suivi d''une modification automatique des fiches d''immobilisation.'+Chr(10)+Chr(13)+'Veuillez corriger manuellement celles-ci.','Avertissement') ;
    MsgBox.Execute(10, '', ''); // Traitement terminé.
  end;

Except
  RollBack ;
  MsgBox.Execute(6, '', ''); // Fin anormale du traitement.
  Liberetobs ;
  EnableControls(Ecran, True) ;
  exit ;
End ;
Liberetobs ;
{$IFDEF EAGLCLIENT}
Fiche.Q.TQ.Open;
{$ELSE}
Fiche.Q.Open ;
{$ENDIF}
  TToolBarButton97(GetControl('BCHERCHE')).Click ;
EnableControls(Ecran, True) ;
End ;

Procedure TOF_CPPRORATATVA.LibereTobs ;
Begin
If TobTTC <> nil then begin TobTTC.Free ; TobTTC:= Nil ; end ;
If TobTVA <> nil then begin TobTVA.Free ; TobTVA:=Nil ; End ;
If TobHT <> nil then begin TobHT.Free ; TobHT:= Nil ; End ;
If TobGen <> nil then begin TobGen.Free ; TobGen:=Nil ; End ;
if TobSec <> nil then begin TobSec.Free ; TobSec:=Nil ; End ;
If TobEcr <> nil then begin TobEcr.Free ; TobEcr:=Nil ; end ;
if TobProrata <> nil then begin TobProrata.Free ; tobProrata:=Nil ; end ;
if TobDev <> nil then begin TobDev.Free ; TobDev:= Nil ; end  ;
End ;

Procedure TOF_CPPRORATATVA.ExerciceOnChange (Sender : TObject);
var Q:Tquery ;
    FinDate : TDateTime ;
    sTmp : string;
// *+*    szRequete : String;
BEGIN
// *+*  szRequete := 'SELECT PA_JOURNAL FROM CPPRORATA WHERE PA_CODE = "' + GetControlText('E_EXERCICE') +'" AND PA_TOUSJAL="X"';
// *+*  if ExisteSQL(szRequete) then
// *+*    JournalBis.Plus := 'AND J_NATUREJAL="ACH" AND J_FERME="-"'
// *+*  else
    {JP 10/06/03 : On n'affiche que les journaux qui correpondent à l'exercice sélectionné, comme ils
                   ont été paramétrés}
    //SG6 14/01/05 Correction pour le controle JOURNAL
    sTmp := ' AND J_JOURNAL IN (SELECT PA_JOURNAL FROM CPPRORATA WHERE PA_CODE = "' + GetControlText('E_EXERCICE') +'")';

    Q:=OpenSql('SELECT PA_JOURNAL FROM CPPRORATA WHERE PA_CODE = "'  + GetControlText('E_EXERCICE') +'"',true);
    if not Q.eof then
    begin
      if (Q.FindField('PA_JOURNAL').AsString = '---') then
      begin
        sTmp := '';
      end;
    end;
    Ferme(Q);
    JournalBis.Plus := 'AND J_NATUREJAL="ACH" AND J_FERME="-"'+sTmp;

  {JP 10/06/03 : Pour la mise à jour de la clause where sur les journaux}
  JournalExit(Sender);

Q := OpenSql('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exercice.Value+'"', True) ;
SetControlText('E_DATECOMPTABLE',DateToStr(Q.Fields[0].AsDateTime)) ;
SetControlText('E_DATECOMPTABLE_',DateToStr(Q.Fields[1].AsDateTime)) ;
SetControlText('DATEGEN',DateToStr(Q.Fields[1].AsDateTime)) ;
FinDate := Q.Fields[1].AsDateTime ;
if FinDate >= StrToDate('01/01/2002') then
  Begin
  Money.Enabled := False ; Money.Value:='E' ;
  end else
  begin
  Money.Enabled := True ; if VH^.TenueEuro then Money.Value:='E' else Money.Value:='F' ;
  end ;
Ferme(Q) ;
END ;

Procedure TOF_CPPRORATATVA.ExerciceOnExit (Sender : TObject);
begin
  //on affiche <<Tous>>
  if (GetControlText('E_JOURNAL')='') or (AncienExo<>GetControlText('E_EXERCICE'))
  then SetControlText('E_JOURNAL','<<Tous>>');
end;

Procedure TOF_CPPRORATATVA.ExerciceOnEnter (Sender : TObject);
begin
  AncienExo := GetControlText('E_EXERCICE');
end;


Procedure TOF_CPPRORATATVA.FListeDblClick (Sender: TObject);
Var sMode : String ;
	{$IFDEF EAGLCLIENT}
		TOBListe,TOBLigne : TOB ;
  {$ENDIF}
begin
inherited;
ListeEcr := Fiche.FListe ;
{$IFDEF EAGLCLIENT}
QListeEcr := Fiche.Q.TQ ;
{$ELSE}
QListeEcr := Fiche.Q ;
{$ENDIF}
if ((QListeEcr.EOF) and (QListeEcr.BOF)) then Exit ;
sMode:=QListeEcr.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then
  // MODE BORDEREAU
  begin
{$IFDEF EAGLCLIENT}
 // A FAIRE Voir SaisBor.PAS
{$ELSE}
  LanceSaisieFolio(QListeEcr,TaConsult)
{$ENDIF}
  end
else
  // MODE PIECE
  begin
    {$IFDEF EAGLCLIENT}
    TOBListe := TOB.Create('Liste Ecriture',nil,-1) ;
    TOBLigne := QListeEcr.Detail[ListeEcr.row-1] ;
    TOB.Create('ECRITURE',TOBListe,-1) ;
    TOBListe.Detail[0].Dupliquer(TOBLigne, False, True);
    TrouveEtLanceSaisie(TOBListe, TaConsult, TOBListe.Detail[0].GetString('E_QUALIFPIECE') ) ;
    FreeAndNil(TOBListe) ;
    {$ELSE}
    TrouveEtLanceSaisie(QListeEcr,TaConsult,QListeEcr.FindField('E_QUALIFPIECE').AsString) ;
    {$ENDIF}
  end;
end;

{Procedure TOF_CPPRORATATVA.DateGenOnExit (Sender : Tobject);
var Q:Tquery ;
begin
Q := OpenSql('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exercice.Value+'"', True) ;
If ((StrToDate(DateGen.Text) < Q.Fields[0].AsDateTime) or (StrToDate(DateGen.Text) > Q.Fields[1].AsDateTime)) then
  Begin
  PgiBox(TraduireMemoire('La date de génération des écritures n''est pas inclue dans l''exercice'),TraduireMemoire(Titre)) ;
  Dategen.SetFocus ;
  end ;
Ferme(Q);
end ;
}

Function TOF_CPPRORATATVA.VerifieTauxDefinitifSaisi : Boolean ;
var Q:TQuery ;
begin
	Result := True;
  Q:=OpenSql('Select PA_TAUX2 FROM CPPRORATA WHERE PA_CODE="'+Exercice.Value+'"',True) ;
  If Q.FindField('PA_TAUX2').AsFloat=0 then begin
    MsgBox.Execute(8, '', ''); // Veuillez saisir le taux de prorata définitif.
    Result := False;
  end;
  Ferme(Q) ;
end ;

procedure TOF_CPPRORATATVA.OnArgument (S : String ) ;
Var
  HelpContext : String ;
  Devise      : THValComboBox;            {FP 08/11/2005 FQ15457}
begin

  // Création des contrôles
  MsgBox := THMsgBox.create(FMenuG) ;
  InitMsgBox;
  QuelProrata:=ReadTokenPipe(S,';') ;
  HelpContext := ReadTokenPipe(S,';') ;
  Ecran.HelpContext := StrToInt(HelpContext) ;

  Fiche := TFMul(Ecran) ;
  ListeEcr :=Fiche.FListe ;
{$IFDEF EAGLCLIENT}
  QListeEcr := Fiche.Q.TQ ;
{$ELSE}
  QListeEcr := Fiche.Q ;
{$ENDIF}
  Dategen := THEdit(GetControl('DATEGEN')) ;
  GenerationAu := THLabel(GetControl('GENERATIONAU')) ;
  //YM 22/09/2005 Dejà fait...
 // QuelProrata:=ReadTokenPipe(S,';') ;
  if QuelProrata='Last' then begin
    Ecran.Caption:='Liste des écritures à régulariser' ;
    SetControlVisible('EcrModifiees', True) ;
    SetControlChecked('EcrModifiees',True) ;
    GenerationAu.Visible :=True ;
    DateGen.Visible:=True ;
    SetControlVisible('Generation',True) ;
    TFMul(Ecran).FNomFiltre := 'CPFICPRORATAR';  // FQ 15452 SBO 25/03/2005
    end
  else begin
    Ecran.Caption:='Liste des écritures à proratiser' ;
    SetControlVisible('EcrModifiees', False) ;
    GenerationAu.Visible :=False ;
    DateGen.Visible:=False ;
    SetControlVisible('Generation',False) ;
  end ;
{JP 06/06/2003 : pour mettre à jour le titre de la fenêtre}
UpdateCaption(Fiche);
Exercice := THValComboBox(GetControl('E_EXERCICE')) ;
JournalBis := THMultiValComboBox(GetControl('E_JOURNAL')) ;
JournalBis.OnExit := JournalExit;

Tva := THValComboBox(GetControl('E_TVA')) ;
Money := THValComboBox(GetControl('MONEY')) ;
Datedeb := THEdit(GetControl('E_DATECOMPTABLE')) ;
DateFin := THEdit(GetControl('E_DATECOMPTABLE_')) ;
General1 := THEdit(GetControl('E_GENERAL')) ;
General2 := THEdit(GetControl('E_GENERAL_')) ;
Pages := TPageControl(GetControl('PAGES')) ;
TabLib := TTabSheet(GetControl('TABLIB')) ;
InitTablesLibres (TabLib) ;
TToolBarButton97(GetControl('BOUVRIR')).OnClick := BOuvrirOnClick ;
Exercice.OnChange := ExerciceOnChange ;
Exercice.OnExit :=ExerciceOnExit;
Exercice.OnEnter := ExerciceOnEnter;
SetControlText('E_EXERCICE',QuelExoDt(Date)) ;
DateDeb.Text := DateToStr(Date) ;
DateFin.Text := DateToStr(Date) ;
{$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick ;
{$ELSE}
  THDBGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick ;
{$ENDIF}

//SG6 12/01/05 FQ 15155 Gestion des écritures de simu
THMultiValComboBox(getcontrol('QUALIFPIECE',true)).OnChange := qualifpiecechange;
THMultiValComboBox(getcontrol('QUALIFPIECE',true)).Plus := 'AND (CO_CODE="N" OR CO_CODE="S")';
LEcr:=TStringList.Create ;
  LEcr.Sorted:=TRUE ;
  LEcr.Duplicates:=dupIgnore ;

  InitValid;

  {b FP 08/11/2005 FQ15457}
  Devise := THValComboBox(getcontrol('E_DEVISE',true));
  Devise.Enabled := False;               {FP 01/02/2006 FQ15457}
  Devise.Value   := V_PGI.DevisePivot;
  {e FP 08/11/2005 FQ15457}

  Inherited ;
  {JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements}
  ComboEtab := THValComboBox(GetControl('E_ETABLISSEMENT'));
  GereEtablissement;
  TFMul(Ecran).OnAfterSelectFiltre := AfterSelectFiltre;;
end ;

// Mémorise le montant HT initial
procedure TOF_CPPRORATATVA.MemoriseHTInitial(TobHT: Tob);
var
  iProrataTVA : Integer;
  dblValeur : Double;
begin
  iProrataTVA := GetParamSocSecur('SO_HTPRORATATVA',0);
  if (iProrataTVA > 0) then begin
    Dec(iProrataTVA);
    dblValeur := TobHT.GetValue('E_DEBIT')-TobHT.GetValue('E_CREDIT');
    TobHT.PutValue('E_LIBREMONTANT'+IntToStr(iProrataTVA), dblValeur);
  end;
end;

{-------------------------------------------------------------------------------------------}
procedure TOF_CPPRORATATVA.JournalExit(Sender : TObject);
{-------------------------------------------------------------------------------------------}
var
  s : string;
const
  w = '(G_NATUREGENE = "CHA" or G_NATUREGENE = "IMO")';
begin
  if GetControlText('E_JOURNAL')='' then SetControlText('E_JOURNAL','<<Tous>>');
  {JP 10/06/03 : Formatage de la clause where pour éviter de traiter tous les journaux et
               non ceux qui correspondent à l'exercice sélectionné}
  s := RetourneClauseIn(GetControlText('E_JOURNAL'));
  if Trim(s) <> '' then
    SetControlText('XX_WHERE','E_JOURNAL in (' + s + ') AND ' + w)
  else
    SetControlText('XX_WHERE', w)
end;

{JP 10/06/06 : Formate la chaine 'ml;po;...' en '"ml", "po", ...
{-------------------------------------------------------------------------------------------}
function TOF_CPPRORATATVA.RetourneClauseIn(Chaine : string) : string;
{-------------------------------------------------------------------------------------------}
var
  s, ch : string;
  n     : Integer;
begin
  if ((UpperCase(Chaine) = '<<TOUS>>') and
      not ((JournalBis.Items.Count = 1) and (UpperCase(JournalBis.Values[0]) = '<<TOUS>>')))
          or(Trim(JournalBis.Text) = '') then
  begin
  { Si on a tous  et pas seulement un seul choix dans la combo qui serait "Tous"
    ou bien rien de sélectionné}
    for n := 0 to JournalBis.Items.Count - 1 do
      s := s + JournalBis.Values[n] + ';';
  end
  else if (UpperCase(Chaine) <> '<<TOUS>>') and (Trim(JournalBis.Text) <> '') then
    s := Chaine
  else
    Exit;

  ch := ReadTokenSt(s);
  while Trim(ch) <> '' do begin
    Result := Result + ' "' + Ch + '", ';
    ch := ReadTokenSt(s);
  end;
  {On enlève éventuellement la dernière virgule}
  if Result <> '' then System.Delete(Result, Length(Result) - 1, 1);
end;

procedure TOF_CPPRORATATVA.InitMsgBox;
begin
  MsgBox.Mess.Add('0;'+Titre+';Confirmez-vous le traitement ?;Q;YN;Y;N;');
  MsgBox.Mess.Add('1;'+Titre+';Aucune ligne sélectionnée. Traitement impossible.;E;O;O;O;');
  MsgBox.Mess.Add('2;'+Titre+';Veuillez paramétrer les taux de prorata de TVA.;E;O;O;O;');
  MsgBox.Mess.Add('3;'+Titre+';Veuillez paramétrer le taux de prorata de TVA définitif pour l''année sélectionnée.;E;O;O;O;');
  MsgBox.Mess.Add('4;'+Titre+';Le taux intermédiaire et définitif sont identiques pour l''année sélectionnée.#10#13Veuillez revoir le paramétrage des taux de prorata de TVA.;E;O;O;O;');
  MsgBox.Mess.Add('5;'+Titre+';ATTENTION : Certaines pièces n''ont pas été proratisées. Lancez le contrôle des factures.;E;O;O;O;');
  MsgBox.Mess.Add('6;'+Titre+';Fin anormale du traitement.;E;O;O;O;');
  MsgBox.Mess.Add('7;'+Titre+';La date de génération des écritures n''est pas inclue dans l''exercice.;E;O;O;O;');
  MsgBox.Mess.Add('8;'+Titre+';Veuillez saisir le taux de prorata définitif.;E;O;O;O;');
  MsgBox.Mess.Add('9;'+Titre+';La date de génération des écritures n''est pas inclue dans l''exercice.;I;O;O;O;');
  MsgBox.Mess.Add('10;'+Titre+';Traitement terminé.;I;O;O;O;');
  MsgBox.Mess.Add('11;'+Titre+';L''écriture n°%% n''est pas en mode pièce.;W;O;O;O;');
  MsgBox.Mess.Add('12;'+Titre+';L''écriture n°%% a déjà été proratisée avec le taux final.;W;O;O;O;');
  MsgBox.Mess.Add('13;'+Titre+';L''écriture n°%% n''est pas de nature Avoir ou Facture.;W;O;O;O;');
  MsgBox.Mess.Add('14;'+Titre+';L''écriture n°%% a plusieurs lignes de TTC.;W;O;O;O;');
  MsgBox.Mess.Add('15;'+Titre+';L''écriture n°%% n''a pas de ligne de TTC.;W;O;O;O;');
  MsgBox.Mess.Add('16;'+Titre+';L''écriture n°%% n''a pas de ligne de HT.;W;O;O;O;');
  MsgBox.Mess.Add('17;'+Titre+';L''écriture n°%% n''a pas de ligne de TVA.;W;O;O;O;');
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 12/01/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet de gérer la condition générée par le
Suite ........ : controle QUALIFPIECE
Mots clefs ... : QUALIFPIECE
*****************************************************************}
procedure TOF_CPPRORATATVA.qualifpiecechange(Sender : TObject);
var
  sSql,tmp1,tmp2 : string;
  qualifp : THMultiValComboBox;
begin

  qualifp := THMultiValComboBox(Sender);
  tmp2 := 'E_QUALIFPIECE';
  if qualifp.Text = '<<Tous>>' then sSql := '(E_QUALIFPIECE = "S" or E_QUALIFPIECE = "N")'
  else TraductionTHMultiValComboBox(qualifp,sSql,tmp1,tmp2,false);
  SetControlText('XX_WHERE5',sSql);

  InitValid;
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 25/08/2005
Modifié le ... :   /  /
Description .. : Case 'écriture validée O/N' uniquement si écritures normales
Suite ........ :
Mots clefs ... : QUALIFPIECE VALID
*****************************************************************}
procedure TOF_CPPRORATATVA.InitValid;
begin
  if (getcontroltext('QUALIFPIECE') = 'N')
  or (getcontroltext('QUALIFPIECE') = 'N;') then
  begin
      SetControlProperty('E_VALIDE', 'VISIBLE', True);
      SetControlProperty('FE__HLabel3', 'VISIBLE', True);
  end
  else
  begin
      SetControlProperty('E_VALIDE', 'VISIBLE', False);
      SetControlProperty('FE__HLabel3', 'VISIBLE', False);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPRORATATVA.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ComboEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      ComboEtab.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      ComboEtab.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(ComboEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if ComboEtab.Value = '' then begin
        {... on affiche l'établissement par défaut
        ComboEtab.Value := VH^.EtablisDefaut;
        {... on active la zone
        ComboEtab.Enabled := True;
      end;}
    end;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPRORATATVA.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(ComboEtab) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;

  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> ComboEtab.Value) then begin
    {... on affiche l'établissement des restrictions}
    ComboEtab.Value := Eta;
    {... on désactive la zone}
    ComboEtab.Enabled := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPRORATATVA.AfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  ControlEtab;
end;

Initialization
  registerclasses ( [ TOF_CPPRORATATVA ] ) ;
end.
