{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 02/03/2004
Description .. : Passage en eAGL
Suite ........ : GCO - 02/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Suite ........ : GCO - 03/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheGene en 2/3 et CWAS
Suite ........ : -> Uniformisation de l'appel à FicheTiers en 2/3 et CWAS
Suite ........ : -> Uniformisation de l'appel à FicheSection en 2/3 et CWAS
Suite ........ : - BTY - 04/06 FQ 17629 - Choix du calcul de la PMVAlue
Mots clefs ... :
*****************************************************************}
unit HalOLE;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls, Forms,
     Menus, Hctrls, DB, Mask, StdCtrls, ExtCtrls, Buttons, ComCtrls, HStatus,
     DBTables, HEnt1, hmsgbox, OLEAuto,ParamSoc,
     UtilSoc,
//     HQuickRP,
{$IFDEF EAGLCLIENT}
     UtileAGL,
     Rubrique_TOM,
     AssistOLE,
{$ELSE}
    {$IFNDEF CCS3}
         BudGene,BudSect,
    {$ENDIF}
    {$IFDEF V530}
          EdtEtat,
    {$ELSE}
          EdtREtat,
    {$ENDIF}
    AssistEF,  // A FAIRE : A Porter Obligatoirement
    Rubrique_TOM,
{$ENDIF}
    CPJournal_TOM,
    CPGeneraux_TOM,
    CPTiers_TOM,
    CPSection_TOM,
    ExcelToEcr,
    MPlayer, HSysMenu,MajTable,Ent1,CalcOLE,HCompte,
    uTofConsEcr,
    UTob, SoldeCpt
    ,GalOutil
{$IFDEF AMORTISSEMENT}
    , Outils
    , ImPlan
    , ImOutGen
    , imPlanInfo
    , AMLISTE_TOF
    , AmInterOLE
{$ENDIF}
;

procedure RegisterHalleyWindow ;

Type
// Automation Server
  THalleyWindow = class(TAutoObject)
  public
    constructor Create ; override;
    destructor Destroy; override;
  private
//    FMainForm : TMainForm ;
    FListe : TStringList ;
    FListeImmo : TStringList;
    FPosListe : Integer ;
    FPosListeImmo : integer;
    function JaiLeDroit : string;
  automated
    Function Cumul(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
    Function CumulMS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String; stRegroupement : string) : Variant ;
    Function Cumul2(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,Qui2,BalSit,WhereSup : String) : Variant ;
    Function Cumul2MS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,Qui2,BalSit,WhereSup : String; stRegroupement : string = '') : Variant ;
    Function CumulDouble(Table,Qui1,Qui2,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
    Function Info(Table,Qui,Quoi,Dollar : String) : Variant ;
    Function GetSQL(SQL : String) : Variant ;
    Function DateCumul(SDate,Quoi : String) : Variant ;
    Function Zoom(Table,Qui : string) : Variant;
    Procedure ChargeListe(Const Quoi : string ) ;
    Function NextListe : string  ;
    Function FinDeListe : Integer  ;
    Function AssistantEF(sFormule : String) : String ;
    Function Constante(CodeCons,CodeEtab,Periode : String) : Variant ;
    Function GetExos : variant ;
    Function GetSocExos : variant ;
    Function GetDevisePivot : Variant ;
    Function GetSociete : Variant ;
    Function GetConnexion : Variant ;
    Function GetCodeSociete : Variant ;
    Function GetNoDossier : Variant ;
    Function GetInfoSociete(Champ : string) : Variant ;
    Function GetBalance(nExo : Integer) : Variant ;
    Function ZoomEtat(Tip,Nat,Modele,sWhere : string) : Variant ;
    Function ConsultationCompte(General,Exercice,ChampsTries : string) : Variant ;
    Function ConsultationCompteDetail(General,Exercice,ChampsTries : string) : Variant ;
    {$IFDEF AMORTISSEMENT}
    Function ConsultationImmos ( General : String ) : Variant ;
    {$ENDIF}
    Function ComptaModified : Variant ;
    Function GetResultat ( CodeExo : String ) : Variant ;
    Function GetPathSociete : Variant ;
    Function GetProgSerie : Variant ;
    Function PlusBesoinOLE : Variant ;
    Function GetListeExos : Variant ;
    (*
    Function Balance(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
    *)
    {IMMOS}
    {$IFDEF AMORTISSEMENT}
    Function GetBaseTaxePro(CpteInf,CpteSup,NatureBien, LieuInf,LieuSup : string) : Variant ;
    Function GetBaseTaxePro2(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup,Eligible : string) : Variant ;
    Function GetAugAcqMut(CpteInf,CpteSup : string) : Variant ;
    Function GetDotationExo(CpteInf,CpteSup, Methode : string) : Variant ;
    Function GetDiminutionSortie(CpteInf,CpteSup : string) : Variant ;
    Function GetDiminutionMutation(CpteInf,CpteSup : string) : Variant ;
    Function GetDiminutionAmort(CpteInf,CpteSup : string) : Variant ;
    function GetListeLieuImmo: Variant;
    function GetValHTImmo( pstCpteInf, pstCpteSup, pstParam : string ) : Variant ;
    function ChargeLesImmos(stType: string; iVersion: integer) : variant;
    function ImmoSuivante(iVersion: integer): Variant;
    procedure ChargeListeImmo(stType : string);
    function GetNextImmo : Variant;
    {$ENDIF}
    function GetUserName : Variant;
    {Excel to PGI}
    function TraiteEcrituresRevision ( LeJournal,LeFolio : String) : Variant ;
    function TraiteEcrituresComptesRevision ( LeJournal,LeFolio : String) : Variant ;
    Function TraiteEcrituresComptesRevisionach ( LeJournal,LeFolio : string ) : Variant ;
    Function TraiteEcrituresComptesRevisionvte ( LeJournal,LeFolio : string ) : Variant ;
    Function TraiteEcrituresComptesRevisionbqe ( LeJournal,LeFolio,Contrepartie : string ) : Variant ;
    Function TraiteEcrituresBalance ( LeJournal,LeFolio,RecupLib : string) : Variant ;

    Function GetListeCollectif (Nature :string) : Variant ;
    Function GetListeTiers (Coll :string) : Variant ;

    function ChargeJournaux: Variant;
    function ChargeVentilType : variant;
    function ChargeSection ( Ax : String ) : variant;

    function GetFolios( stJournal: string ): Variant;
    function GetSeriaRevision : Variant;
    function GetListeSituation ( TypeCum, Exo : string ) : Variant;
    function GetBalanceSituation( TypeCum, CodeBal: string): Variant;
    function GetResultatSituation(CodeBal: string): Variant;
    function ComptaInitialisee : variant;
    procedure MinimiseCompta;
    procedure MaximiseCompta;
    function GetInfoCRE: Variant;
    procedure MajSoldeComptes;
    function GetBalanceCompte(nExo: integer; General, Auxiliaire, Balsit: string): variant;
    Function GetNatureJrl (Journal :string) : Variant ;
    function ChargeEtablissement : variant;
    Function GetContrepartieJrl (Journal :string) : Variant ;
    Function GetModesaisieJrl (Journal :string) : Variant ;
    function GetCreditBails(LDateDeb: variant): Variant;
    function GetImmosFIMVT(LDateDeb: variant): Variant;
    function GetJustImmos(LDateDeb: variant): Variant;
    function GetJustSorties(LDateDeb: variant): Variant;
    function GetPREPDADS(lDateDeb, lDateFin: variant): Variant;
    function GetPROVCPRTT(lDateArrete: variant): Variant;

    // GCO - 19/01/2006
    function EstTvaActivee : Variant;
    function GetMontantTVA( vStCodeTva : string; vDateDebut, vDateFin : TDateTime) : Variant;
    // Fin GCO - 19/01/2006

  end;

implementation

uses Registry;

const
   cSeparateur = ';';
   cRetourChariot = chr(13);

{================================ OLE =====================================}
procedure RegisterHalleyWindow ;
var
  Reg : TRegistry;
const
  AutoClassInfo: TAutoClassInfo = (
    AutoClass: THalleyWindow;
    ProgID: 'Halley.HalleyWindow';
    ClassID: '{23ED1F61-1C6D-11D0-B66C-00AA0035AECD}';
    Description: 'Serveur Halley';
    Instancing: acMultiInstance);
begin
  // FQ 13538
  Reg := TRegistry.Create;
  Reg.RootKey := HKEY_CLASSES_ROOT;
  Reg.DeleteKey('CLSID\{23ED1F61-1C6D-11D0-B66C-00AA0035AECD}');
  Reg.DeleteKey('Halley.HalleyWindow');
  Reg.Free;

  Automation.RegisterClass(AutoClassInfo);
end;


constructor THalleyWindow.Create ;
BEGIN
Inherited create ;
FListe:=TStringList.Create ;
FListeImmo := TStringList.Create;
FPosListe:=0 ;
FPosListeImmo := 0;
END ;

destructor THalleyWindow.Destroy ;
BEGIN
FListe.Free ;
FListeImmo.Free;
Inherited Destroy ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ROHAULT Régis
Créé le ...... : 17/05/2005
Modifié le ... : 17/05/2005
Description .. : Rajout dun verbe OLE MultiSoc (...MS)
Suite ........ : avec un paramètre supplémentaire stRegroupement
Suite ........ : qui est le regroupement des bases.
Mots clefs ... : REGROUPEMENT;MULTIDOSSIER,MULTISOCIETE
*****************************************************************}
Function THalleyWindow.CumulMS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String ; stRegroupement : string) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;

Result := 'Opération non autorisée. Droits insuffisants';
if not ExJaiLeDroitConcept(ccOLE,False) then exit;
cumulMS:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
If Table='RUBRIQUE' Then BEGIN
  if not ExisteSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_RUBRIQUE ="'+Qui+'"') then
    Qui := '@'+Qui;
end;
CumulMS:=Get_CumulMS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif, stRegroupement) ;
END ;

Function THalleyWindow.Cumul2MS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,Qui2,BalSit,WhereSup : String; stRegroupement : string = '') : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
cumul2MS:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Cumul2MS:=Get_Cumul2MS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,Qui2,BalSit,WhereSup,'',NULL, stRegroupement) ;
END ;




Function THalleyWindow.GetSocExos : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=GetCodeSociete+#10+GetExos ;
VH^.CPAppelOLE:=True ;
END ;

Function THalleyWindow.GetExos : Variant ;
var s : string ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
s:='' ;
if VH^.Encours.Code<>'' then
   begin
   s:=VH^.Encours.Code+';' ;
   s:=s+FormatDateTime('ddmmyyyy',VH^.Encours.Deb) + ';' ;
   s:=s+FormatDateTime('ddmmyyyy',VH^.Encours.Fin) + ';' ;
   s:=s+IntToStr(VH.Encours.NombrePeriode) ;
   end else s:=s+'---' ;
s:=s+#10 ;
if VH^.Precedent.Code<>'' then
   begin
   s:=s+VH^.Precedent.Code+';' ;
   s:=s+FormatDateTime('ddmmyyyy',VH^.Precedent.Deb) + ';' ;
   s:=s+FormatDateTime('ddmmyyyy',VH^.Precedent.Fin) + ';' ;
   s:=s+IntToStr(VH.Precedent.NombrePeriode) ;
   end else s:=s+'---' ;
s:=s+#10 ;
if VH^.Suivant.Code<>'' then
   begin
   s:=s+VH^.Suivant.Code+';' ;
   s:=s+FormatDateTime('ddmmyyyy',VH^.Suivant.Deb) + ';' ;
   s:=s+FormatDateTime('ddmmyyyy',VH^.Suivant.Fin) + ';' ;
   s:=s+IntToStr(VH.Suivant.NombrePeriode) ;
   end else s:=s+'---' ;
result:=s ;
VH^.CPAppelOLE:=True ;
end ;

Function THalleyWindow.GetListeExos : Variant ;
var s : string ;
    Q : TQuery ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
s:='' ;
Q:=OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE) ;
While Not Q.Eof Do
  BEGIN
  s:=s+Q.FindField('EX_EXERCICE').AsString+';' ;
  s:=s+FormatDateTime('ddmmyyyy',Q.FindField('EX_DATEDEBUT').AsDateTime) + ';' ;
  s:=s+FormatDateTime('ddmmyyyy',Q.FindField('EX_DATEFIN').AsDateTime) + ';' ;
//  s:=s+Q.FindField('EX_EXERCICE').AsString+';' ;
  s:=s+#10 ;
  Q.Next ;
  END ;
Ferme(Q) ;
Result:=s ;
VH^.CPAppelOLE:=True ;
end ;

Function THalleyWindow.PlusBesoinOLE : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=True ;
VH^.CPAppelOLE:=False ;
END ;

Function THalleyWindow.GetProgSerie : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
if V_PGI.LaSerie=S5 then Result:='CCS5' else Result:='CCS7' ;
END ;

Function THalleyWindow.GetPathSociete : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
if V_PGI.Driver in [dbINTRBASE,dbMSSQL,dbMSACCESS,dbPARADOX,dbSQLBASE]
   then
   {$IFNDEF EAGLCLIENT}
    Result:=ExtractFilePath(GetDBPathName(TRUE))
   {$ELSE}
    Result:=ExtractFilePath(V_PGI.GetNamePath)
   {$ENDIF}
   else Result:='' ;
END ;

Function THalleyWindow.GetDevisePivot : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
result:=V_PGI.DevisePivot ;
end ;

Function THalleyWindow.GetSociete : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
result:=V_PGI.NomSociete ;
end ;

Function THalleyWindow.GetConnexion : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result:=V_PGI.CurrentAlias ;
end ;

Function THalleyWindow.GetResultat ( CodeExo : String ) : Variant ;
Var Q : TQuery ;
    SQL,LeExo : String ;
    XPro,XCha : Double ;
    BPrec,OkOk : Boolean ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=0 ; XPro:=0 ; XCha:=0 ; BPrec:=FALSE ;
if CodeExo='0' then LeExo:=VH^.Encours.Code else
 if CodeExo='-1' then
  BEGIN
    LeExo:=VH^.Precedent.Code ;
//         bPrec:=TRUE ;
//          CA - 27/02/2002 - Ne pas prendre les infos d'histobal si il y a des écritures sur N-1
    bPrec := not (ExisteSQL ('SELECT E_EXERCICE FROM ECRITURE WHERE E_EXERCICE="'+VH^.Precedent.Code+'"'));

  END else
  if ((CodeExo='1') or (CodeExo='+1')) then LeExo:=VH^.Suivant.Code else Exit ;
OkOk:=TRUE ;
If BPrec Then
  BEGIN
  Q:=OpenSQL('SELECT HB_COMPTE1, G_LIBELLE, HB_DEBIT, HB_CREDIT FROM HISTOBAL LEFT JOIN GENERAUX ON HB_COMPTE1=G_GENERAL WHERE HB_TYPEBAL="N1C" OR HB_TYPEBAL="N1A"', TRUE) ;
  OkOk:=Q.EOF ;
  Ferme(Q) ;
  END ;
// Produits
If OkOk Then
  BEGIN
  SQL:='SELECT SUM(E_CREDIT-E_DEBIT) FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL '
      +'WHERE E_EXERCICE="'+LeExo+'" AND E_ECRANOUVEAU="N" AND E_QUALIFPIECE="N" AND G_NATUREGENE="PRO"' ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then XPro:=Q.Fields[0].AsFloat ;
  Ferme(Q) ;
  // Charges
  SQL:='SELECT SUM(E_DEBIT-E_CREDIT) FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL '
      +'WHERE E_EXERCICE="'+LeExo+'" AND E_ECRANOUVEAU="N" AND E_QUALIFPIECE="N" AND G_NATUREGENE="CHA"' ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then XCha:=Q.Fields[0].AsFloat ;
  Ferme(Q) ;
  END Else
  BEGIN
  SQL:='SELECT SUM(HB_CREDIT-HB_DEBIT) FROM HISTOBAL LEFT JOIN GENERAUX ON HB_COMPTE1=G_GENERAL '
      +'WHERE HB_EXERCICE="'+LeExo+'" AND (HB_TYPEBAL="N1C" OR HB_TYPEBAL="N1A") AND G_NATUREGENE="PRO"' ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then XPro:=Q.Fields[0].AsFloat ;
  Ferme(Q) ;
  // Charges
  SQL:='SELECT SUM(HB_DEBIT-HB_CREDIT) FROM HISTOBAL LEFT JOIN GENERAUX ON HB_COMPTE1=G_GENERAL '
      +'WHERE HB_EXERCICE="'+LeExo+'" AND (HB_TYPEBAL="N1C" OR HB_TYPEBAL="N1A") AND G_NATUREGENE="CHA"' ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then XCha:=Q.Fields[0].AsFloat ;
  Ferme(Q) ;
  END ;
Result:=Arrondi(XPro-XCha,V_PGI.OkDecV) ;
END ;

Function THalleyWindow.GetCodeSociete : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;

Result:=V_PGI.CodeSociete ;
VH^.CPAppelOLE:=True ;
end ;

Function THalleyWindow.GetNoDossier : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;

Result:=V_PGI.NoDossier ;
VH^.CPAppelOLE:=True ;
end ;

Function THalleyWindow.GetSQL(SQL : String) : Variant ;
Var Q : tQuery ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Q:=OpenSQL(SQL,TRUE) ;
If Not Q.Eof Then Result:=Q.Fields[0].AsVariant Else Result:='#Erreur : la requète ne renvoie aucun résultat' ;
Ferme(Q) ;
end ;

Function THalleyWindow.GetInfoSociete(Champ : string) : Variant ;
var Q : TQuery ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
if Champ='SO_SOCIETE' then BEGIN Result:=GetSociete ; Exit ; END ;
try
{$IFDEF SPEC302}
   Q:=OpenSQL('SELECT '+Champ+' from SOCIETE',True) ;
   Result:=Q.Fields[0].AsVariant ;
   Ferme(Q) ;
{$ELSE}
   Q:=OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="'+Champ+'"',True) ;
   Result:=Q.Fields[0].AsVariant ;
   Ferme(Q) ;
{$ENDIF}
   except
   result:='#Erreur : Champ société incorrect' ;
   end ;
end ;

Function VirePV(St : String) : String ;
BEGIN
Result:=FindEtReplace(St,';',' ',TRUE) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/06/2003
Modifié le ... : 16/06/2003
Description .. : Ce verbe OLE renvoie les cumuls des comptes généraux
Suite ........ : ou auxiliaires stockés dans la base de données.
Suite ........ : Paramètres :
Suite ........ : - nExo : exercice à prendre en compte ( 0 : encours; -1
Suite ........ : précédent; 1 : suivant; 2 exercice de référence ).
Suite ........ : - General : compte général à traiter
Suite ........ : - Auxiliaire : compte auxiliaire à traiter.
Suite ........ : - Balsit : nom de la balance de situation à interroger
Mots clefs ... :
*****************************************************************}
function THalleyWindow.GetBalanceCompte ( nExo : integer; General : string ; Auxiliaire : string ; Balsit : string ) : variant;
var Q : TQuery ; bPrec, bN1C : Boolean ;
    s,sChampDEB,sChampCRE,sExo : string ;
    TotDeb,TotCre,SoldeDeb,SoldeCre : Double ;
    stQuery : string;
begin
  Result := JaiLeDroit;
  if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;

  bPrec := False;
  if (nExo = 2 ) then
  begin
    if VH^.CPExoRef.Code = VH^.Encours.Code then nExo :=0
    else if VH^.CPExoRef.Code = VH^.Suivant.Code then nExo := 1
    else if VH^.CPExoRef.Code = VH^.Precedent.Code then nExo := -1
    else nExo :=0;
  end;
  case nExo of
    0 :
      begin
        if Auxiliaire='' then
        begin
          sChampDEB:='G_TOTDEBE' ;
          sChampCRE:='G_TOTCREE' ;
        end else
        begin
          sChampDEB:='T_TOTDEBE' ;
          sChampCRE:='T_TOTCREE' ;
        end;
        sExo:='N' ;
      end ;
    -1 :
      begin
        if Auxiliaire='' then
        begin
          sChampDEB:='G_TOTDEBP' ;
          sChampCRE:='G_TOTCREP' ;
        end else
        begin
          sChampDEB:='T_TOTDEBP' ;
          sChampCRE:='T_TOTCREP' ;
        end;
        sExo:='N-' ;
        bPrec := not (ExisteSQL ('SELECT E_EXERCICE FROM ECRITURE WHERE E_EXERCICE="'+VH^.Precedent.Code+'"'));
      end ;
     1 :
      begin
        if Auxiliaire='' then
        begin
          sChampDEB:='G_TOTDEBS' ;
          sChampCRE:='G_TOTCRES' ;
        end else
        begin
          sChampDEB:='T_TOTDEBS' ;
          sChampCRE:='T_TOTCRES' ;
        end;
        sExo:='N+' ;
      end ;
  end ;
  if BalSit <> '' then
  begin
    Q := OpenSQL (' SELECT BSE_CODEBAL,BSE_COMPTE1,BSE_COMPTE2,BSE_DEBIT,BSE_CREDIT FROM CBALSITECR WHERE BSE_CODEBAL="'+
                      BalSit+'" AND BSE_COMPTE1="'+General+'" AND BSE_COMPTE2="'+Auxiliaire+'"',True );
    if not Q.Eof then
    begin
      TotDeb:=Q.FindField('BSE_DEBIT').AsFloat ;
      TotCre:=Q.FindField('BSE_CREDIT').AsFloat ;
      SoldeDeb:=0 ; SoldeCre:=0 ;
      if TotDeb-TotCre>0 then SoldeDeb:=Arrondi(TotDeb-TotCre, V_PGI.OkDecV)
                         else SoldeCre:=Arrondi(TotCre-TotDeb, V_PGI.OkDecV) ;
      s:=Strfpoint(TotDeb) + ';' + Strfpoint(TotCre) +';' ;
      s:=s+Strfpoint(SoldeDeb) + ';' + Strfpoint(SoldeCre) ;
      s:=s+#10 ;
      Result := s;
    end
    else Result := '';
    exit;
  end;
  if bPrec then
  begin
    Q:=OpenSQL('SELECT HB_COMPTE1, HB_COMPTE2, HB_DEBIT, HB_CREDIT FROM HISTOBAL WHERE (HB_TYPEBAL="N1C" OR HB_TYPEBAL="N1A")'+
                ' AND (HB_COMPTE1="'+General+'" AND HB_COMPTE2="'+Auxiliaire+'") ORDER BY HB_TYPEBAL,HB_COMPTE1,HB_COMPTE2', TRUE) ;  //CA - 27/11/2001
    bN1C:=not Q.EOF ;
    if not Q.EOF then
    begin
      s:=s+Strfpoint(0) + ';' + Strfpoint(0) +';' ;
      s:=Strfpoint(Q.FindField('HB_DEBIT').AsFloat) + ';' + Strfpoint(Q.FindField('HB_CREDIT').AsFloat) ;
      s:=s+#10 ;
    end ;
    Ferme(Q) ;
    if bN1C then
    begin
      result:=s ;
    end else Result := '';
    exit;
  end ;
  if Auxiliaire = '' then
  begin
    stQuery := ' SELECT G_GENERAL,G_TOTDEBE,G_TOTCREE,G_TOTDEBP,G_TOTCREP,G_TOTDEBS,G_TOTCRES ';
    stQuery := stQuery + ' FROM GENERAUX ';
    stQuery := stQuery + ' WHERE G_GENERAL="'+General+'"';
  end
  else
  begin
    stQuery := ' SELECT T_AUXILIAIRE, T_TOTDEBE,T_TOTCREE,T_TOTDEBP,T_TOTCREP,T_TOTDEBS,T_TOTCRES ';
    stQuery := stQuery + ' FROM TIERS ';
    stQuery := stQuery + ' WHERE T_AUXILIAIRE="'+Auxiliaire+'"';
  end;
  Q := OpenSQL ( stQuery, False );
  if not Q.EOF then
  begin
    TotDeb:=Q.FindField(sChampDEB).AsFloat ;
    TotCre:=Q.FindField(sChampCRE).AsFloat ;
    SoldeDeb:=0 ; SoldeCre:=0 ;
    if TotDeb-TotCre>0 then SoldeDeb:=Arrondi(TotDeb-TotCre, V_PGI.OkDecV)
                       else SoldeCre:=Arrondi(TotCre-TotDeb, V_PGI.OkDecV) ;
    s:=Strfpoint(TotDeb) + ';' + Strfpoint(TotCre) +';' ;
    s:=s+Strfpoint(SoldeDeb) + ';' + Strfpoint(SoldeCre) ;
    s:=s+#10 ;
  end ;
  Ferme(Q) ;
  result:=s ;
end;

Function THalleyWindow.GetBalance(nExo : Integer) : Variant ;
var Q : TQuery ; bPrec, bN1C : Boolean ;
    s,sChampDEB,sChampCRE,sExo : string ;
    TotDeb,TotCre,SoldeDeb,SoldeCre : Double ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
s:='' ; bPrec:=FALSE ;
if (nExo = 2 ) then // CA - 26/12/2001 - Gestion de l'exercice de référence
begin
  if VH^.CPExoRef.Code = VH^.Encours.Code then nExo :=0
  else if VH^.CPExoRef.Code = VH^.Suivant.Code then nExo := 1
  else if VH^.CPExoRef.Code = VH^.Precedent.Code then nExo := -1
  else nExo :=0;
end;
case nExo of
     0 : begin
         sChampDEB:='G_TOTDEBE' ;
         sChampCRE:='G_TOTCREE' ;
         sExo:='N' ;
         end ;
    -1 : begin
         sChampDEB:='G_TOTDEBP' ;
         sChampCRE:='G_TOTCREP' ;
         sExo:='N-' ;
//         bPrec:=TRUE ;
//          CA - 01/02/2002 - Ne pas prendre les infos d'histobal si il y a des écritures sur N-1
         bPrec := not (ExisteSQL ('SELECT E_EXERCICE FROM ECRITURE WHERE E_EXERCICE="'+VH^.Precedent.Code+'"'));
         end ;
     1 : begin
         sChampDEB:='G_TOTDEBS' ;
         sChampCRE:='G_TOTCRES' ;
         sExo:='N+' ;
         end ;
     end ;
if bPrec then
  begin
  Q:=OpenSQL('SELECT HB_COMPTE1, G_LIBELLE, HB_DEBIT, HB_CREDIT FROM HISTOBAL LEFT JOIN GENERAUX ON HB_COMPTE1=G_GENERAL WHERE HB_TYPEBAL="N1C" OR HB_TYPEBAL="N1A"'+
                ' ORDER BY HB_TYPEBAL,HB_COMPTE1,HB_COMPTE2', TRUE) ;  //CA - 27/11/2001
  bN1C:=not Q.EOF ;
  while not Q.EOF do
  begin
    s:=s+VirePV(Q.FindField('HB_COMPTE1').AsString) + ';' + VirePV(Q.FindField('G_LIBELLE').AsString) + ';' ;
    s:=s+Strfpoint(0) + ';' + Strfpoint(0) +';' ;
    s:=s+Strfpoint(Q.FindField('HB_DEBIT').AsFloat) + ';' + Strfpoint(Q.FindField('HB_CREDIT').AsFloat) ;
    s:=s+#10 ;
    Q.Next ;
  end ;
  Ferme(Q) ;
  if bN1C then
  begin
    result:=s ;
    MarquerPublifi(FALSE) ;
  end else Result := '';
  exit;
end ;
Q:=OpenSQL('Select G_GENERAL,G_LIBELLE,G_COLLECTIF,G_TOTDEBE,G_TOTCREE,G_TOTDEBP,G_TOTCREP,G_TOTDEBS,G_TOTCRES FROM GENERAUX ORDER BY G_GENERAL',False) ;
While not Q.EOF do
   begin
   s:=s+VirePV(Q.FindField('G_GENERAL').AsString) + ';' + VirePV(Q.FindField('G_LIBELLE').AsString) + ';' ;
   TotDeb:=Q.FindField(sChampDEB).AsFloat ;
   TotCre:=Q.FindField(sChampCRE).AsFloat ;
   SoldeDeb:=0 ; SoldeCre:=0 ;
   if Q.FindField('G_COLLECTIF').AsString='X' then
      begin
      Get_CumulColl(Q.FindField('G_GENERAL').AsString,sExo,SoldeDeb,SoldeCre) ;
      end else
      begin
      if TotDeb-TotCre>0 then SoldeDeb:=Arrondi(TotDeb-TotCre, V_PGI.OkDecV)
                         else SoldeCre:=Arrondi(TotCre-TotDeb, V_PGI.OkDecV) ;
      end ;
   s:=s+Strfpoint(TotDeb) + ';' + Strfpoint(TotCre) +';' ;
   s:=s+Strfpoint(SoldeDeb) + ';' + Strfpoint(SoldeCre) ;
   s:=s+#10 ;
   Q.Next ;
   end ;
Ferme(Q) ;
result:=s ;
MarquerPublifi(False);
end ;
(*
Function THalleyWindow.GetPeriodeBalance(sExo : String ; Situ : Boolean) : Variant ;
var Q : TQuery ; bPrec, bN1C : Boolean ;
    s : string ;
    TotDeb,TotCre,SoldeDeb,SoldeCre : Double ;
begin
s:='' ;
if Situ then
  begin
  Q:=OpenSQL('SELECT HB_COMPTE1, G_LIBELLE, HB_DEBIT, HB_CREDIT FROM HISTOBAL LEFT JOIN GENERAUX ON HB_COMPTE1=G_GENERAL WHERE HB_TYPEBAL="N1C" OR HB_TYPEBAL="N1A"', TRUE) ;
  bN1C:=Q.EOF ;
  while not Q.EOF do
    begin
    s:=s+VirePV(Q.FindField('HB_COMPTE1').AsString) + ';' + VirePV(Q.FindField('G_LIBELLE').AsString) + ';' ;
    s:=s+Strfpoint(0) + ';' + Strfpoint(0) +';' ;
    s:=s+Strfpoint(Q.FindField('HB_DEBIT').AsFloat) + ';' + Strfpoint(Q.FindField('HB_CREDIT').AsFloat) ;
    s:=s+#10 ;
    Q.Next ;
    end ;
  Ferme(Q) ;
  if bN1C then begin result:=s ; MarquerPublifi(FALSE) ; end ;
  end Else
  BEGIN
  Q:=OpenSQL('Select G_GENERAL,G_LIBELLE,G_COLLECTIF,G_TOTDEBE,G_TOTCREE,G_TOTDEBP,G_TOTCREP,G_TOTDEBS,G_TOTCRES FROM GENERAUX ORDER BY G_GENERAL',False) ;
  While not Q.EOF do
     begin
     s:=s+VirePV(Q.FindField('G_GENERAL').AsString) + ';' + VirePV(Q.FindField('G_LIBELLE').AsString) + ';' ;
     TotDeb:=Q.FindField(sChampDEB).AsFloat ;
     TotCre:=Q.FindField(sChampCRE).AsFloat ;
     SoldeDeb:=0 ; SoldeCre:=0 ;
     if Q.FindField('G_COLLECTIF').AsString='X' then
        begin
        Get_CumulColl(Q.FindField('G_GENERAL').AsString,sExo,SoldeDeb,SoldeCre) ;
        end else
        begin
        if TotDeb-TotCre>0 then SoldeDeb:=TotDeb-TotCre
                           else SoldeCre:=TotCre-TotDeb ;
        end ;
     s:=s+Strfpoint(TotDeb) + ';' + Strfpoint(TotCre) +';' ;
     s:=s+Strfpoint(SoldeDeb) + ';' + Strfpoint(SoldeCre) ;
     s:=s+#10 ;
     Q.Next ;
     end ;
  Ferme(Q) ;
  END ;
result:=s ;
MarquerPublifi(False) ;
end ;
*)
Function THalleyWindow.ZoomEtat(Tip,Nat,Modele,sWhere : string) : Variant ;
begin
InitZoomOLE ;
MakeZoomOle ( V_PGI.ZoomOleHwnd );
Result:=LanceEtat(Tip,Nat,Modele,True,False,False,nil,sWhere,'',False) ;
FinZoomOLE ;
end ;

Function THalleyWindow.ComptaModified : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  Result := GetParamSocSecur('SO_ACHARGERPUBLIFI', False) ;
end ;

{$IFDEF AMORTISSEMENT}
Function THalleyWindow.ConsultationImmos ( General : string ) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=True ;
InitZoomOLE ;
MakeZoomOle ( V_PGI.ZoomOleHwnd );
AMLanceFiche_ListeDesImmobilisations ( General , False, taConsult );
FinZoomOLE ;
END ;
{$ENDIF}

Function THalleyWindow.ConsultationCompte(General,Exercice,ChampsTries : string) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=True ;
InitZoomOLE ;
MakeZoomOle ( V_PGI.ZoomOleHwnd );
  OperationsSurComptes(General,Exercice,ChampsTries) ;
FinZoomOLE ;
END ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/09/2004
Modifié le ... :   /  /
Description .. : Renvoi dans le Variant le détail des écritures de la
Suite ........ : Consultation des écritures
Mots clefs ... :
*****************************************************************}
Function THalleyWindow.ConsultationCompteDetail(General,Exercice,ChampsTries : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  InitZoomOLE ;
  MakeZoomOle ( V_PGI.ZoomOleHwnd );
  Result := OperationsSurComptesDetail(General, Exercice, ChampsTries) ;
  FinZoomOLE ;
end ;
////////////////////////////////////////////////////////////////////////////////

Procedure THalleyWindow.ChargeListe(Const Quoi : string ) ;
Var Q : TQuery ;
BEGIN
if JaiLeDroit <> '' then exit;
FListe.Clear ; FPosListe:=0 ;
if Not V_PGI.OKOuvert then Exit ;
if UpperCase(Quoi)='GENERAL' then Q:=OpenSQL('Select G_GENERAL From Generaux Order by G_GENERAL',TRUE) else
if UpperCase(Quoi)='COLLECTIF' then Q:=OpenSQL('Select G_GENERAL From Generaux Where G_COLLECTIF="X" Order by G_GENERAL',TRUE) else
if UpperCase(Quoi)='CHARGE' then Q:=OpenSQL('Select G_GENERAL From Generaux Where G_NATUREGENE="CHA" Order by G_GENERAL',TRUE) else
if UpperCase(Quoi)='PRODUIT' then Q:=OpenSQL('Select G_GENERAL From Generaux Where G_NATUREGENE="PRO" Order by G_GENERAL',TRUE) else
if UpperCase(Quoi)='BANQUE' then Q:=OpenSQL('Select G_GENERAL From Generaux Where G_NATUREGENE="BQE" or G_NATUREGENE="CAI" Order by G_GENERAL',TRUE) else
if UpperCase(Quoi)='TIERS' then Q:=OpenSQL('Select T_AUXILIAIRE From TIERS Order by T_AUXILIAIRE',TRUE) else
if UpperCase(Quoi)='CLIENT' then Q:=OpenSQL('Select T_AUXILIAIRE From TIERS Where T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD" Order by T_AUXILIAIRE',TRUE) else
if UpperCase(Quoi)='FOURNISSEUR' then Q:=OpenSQL('Select T_AUXILIAIRE From TIERS Where T_NATUREAUXI="FOU" or T_NATUREAUXI="AUC" Order by T_AUXILIAIRE',TRUE) else
if UpperCase(Quoi)='SALARIE' then Q:=OpenSQL('Select T_AUXILIAIRE From TIERS Where T_NATUREAUXI="SAL" Order by T_AUXILIAIRE',TRUE) else
// ajout me 16-06-2004 pour ventilation type
if UpperCase(Quoi)='VENTIL' then Q:=OpenSql('Select CC_CODE from CHOIXCOD Where CC_TYPE="VTY" ', TRUE) else
Exit ;
While Not Q.Eof do
   BEGIN
   FListe.Add(Q.Fields[0].AsString) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

(*Function THalleyWindow.AssistantEF(sFormule : String) : String ;
BEGIN
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
InitZoomOLE ;
MakeZoomOle(V_PGI.ZoomOleHwnd );
{$IFDEF EAGLCLIENT}
Result := 'Non implémenté en CWAS'; // A FAIRE
{$ELSE}
Result:=LanceAssistEF(sFormule) ;
{$ENDIF}
FinZoomOLE ;
END ;
*)

Function THalleyWindow.AssistantEF(sFormule : String) : String ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
InitZoomOLE ;
MakeZoomOle(V_PGI.ZoomOleHwnd );
{$IFDEF EAGLCLIENT}
Result:=LanceAssistOLE(sFormule) ;
{$ELSE}
Result:=LanceAssistEF(sFormule) ;
{$ENDIF}
FinZoomOLE ;
END ;

Function THalleyWindow.NextListe : string  ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
NextListe:='' ;
if Not V_PGI.OKOuvert then Exit ;
if FPosListe>=FListe.Count then exit ;
NextListe:=FListe[FPosListe] ;
Inc(FPosListe) ;
END ;

Function THalleyWindow.FinDeListe : Integer  ;
BEGIN
FinDeListe:=-1 ;
if JaiLeDroit <> '' then exit;
if Not V_PGI.OKOuvert then Exit ;
FinDeListe:=-ord((FPosListe>=FListe.Count)) ;
END ;

Function THalleyWindow.Constante(CodeCons,CodeEtab,Periode : String) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
Constante:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Constante:=Get_Constante(CodeCons,CodeEtab,Periode) ;
END ;

Function THalleyWindow.Cumul(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
cumul:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
If Table='RUBRIQUE' Then BEGIN
  if not ExisteSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_RUBRIQUE ="'+Qui+'"') then
    Qui := '@'+Qui;
end;
Cumul:=Get_Cumul(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif) ;
END ;

Function THalleyWindow.Cumul2(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,Qui2,BalSit,WhereSup : String) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
cumul2:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Cumul2:=Get_Cumul2(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,Qui2,BalSit,WhereSup,'',NULL) ;
END ;

Function THalleyWindow.CumulDouble(Table,Qui1,Qui2,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
cumulDouble:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
CumulDouble:=Get_CumulDouble(Table,Qui1,Qui2,AvecAno,Etab,Devi,SDate,Collectif) ;
//CumulDouble:=0 ;
END ;

Function THalleyWindow.Info(Table,Qui,Quoi,Dollar : String) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
Info:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Info:=Get_Info(Table,Qui,Quoi,Dollar) ;
END ;

Function THalleyWindow.DateCumul(SDate,Quoi : String) : Variant ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
DateCumul:='#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result:=Get_DateCumul(SDate,Quoi) ;
END ;

Function THalleyWindow.Zoom(Table,Qui : string) : Variant;
Var Sql,St1 : String ;
    Err,Axe : String ;
    fb : TFichierBase ;
BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
St1:='' ; Axe:='' ; fb:=fbGene ; Err:='' ;
if(Table='') then BEGIN Result:='#ERREUR:Table non renseignée' ; Exit ; END ;
if(Qui='') then BEGIN Result:='#ERREUR:Code recherché non renseigné' ; Exit ; END ;
If Table='GENERAUX' Then BEGIN St1:='G_GENERAL' ; fb:=fbGene ; Qui:=BourreLaDonc(Qui,fb) ; END Else
If Table='BUDGENE' Then BEGIN St1:='BG_BUDGENE' ; fb:=fbBudGen ; END Else
If Table='RUBRIQUE' Then BEGIN St1:='RB_RUBRIQUE' ; fb:=fbRubrique ; END Else
If Table='TIERS'    Then BEGIN St1:='T_AUXILIAIRE' ; fb:=fbAux ; Qui:=BourreLaDonc(Qui,fb) ; END Else
If Table='SECTION'  Then BEGIN St1:='S_SECTION' ; Axe:=Copy(Qui,1,2) ; Qui:=Copy(Qui,3,Length(Qui)-2) ; fb:=TFichierBase(Ord(Axe[2])-49) ; Qui:=BourreLaDonc(Qui,fb) ; END Else
If Table='BUDSECT'  Then BEGIN St1:='BS_BUDSECT' ; Axe:=Copy(Qui,1,2) ; Qui:=Copy(Qui,3,Length(Qui)-2) ; fb:=TFichierBase(Ord(fbBudSec1)+Ord(Axe[2])-49) ; END Else
If Table='JOURNAL'  Then St1:='J_JOURNAL' Else
If Table='SOCIETE'  Then BEGIN ST1:='SO_SOCIETE' ; Qui:=V_PGI.CodeSociete END
                    else BEGIN Result:='#ERREUR:Table inexistante' ; Exit ; END ;
Sql:='Select '+St1+' From '+Table+' Where '+St1+'="'+Qui+'"' ;
If Fb in [fbAxe1..fbAxe5] Then Sql:= Sql + 'AND S_AXE="'+AXE+'"' else
If Fb in [fbBudSec1..fbBudSec5] Then Sql := Sql + 'AND BS_AXE="'+AXE+'"';
If Table='RUBRIQUE' Then BEGIN
  if not ExisteSQL(Sql) then begin
    Qui := '@'+Qui;
    Sql:='Select '+St1+' From '+Table+' Where '+St1+'="'+Qui+'"' ;
  end;
end;

if not ExisteSQL(Sql) then begin
  Err:='#ERREUR:Enregistrement non trouvé';
  END
else  BEGIN
Result := JaiLeDroit;
if Result <> '' then exit;
  InitZoomOLE ;
  MakeZoomOle ( V_PGI.ZoomOleHwnd );
  Case fb of
    fbGene : FicheGene(Nil,'',Qui,taConsult,0);
    fbAux  : FicheTiers(Nil,'',Qui,taConsult,0);
    fbAxe1..fbAxe5 : FicheSection(Nil,Axe,Qui,taConsult,0);
    {$IFDEF EAGLCLIENT}
      {$IFNDEF CCS3}
      fbBudGen : Err:='#ERREUR:Non implémenté en CWAS';  // A FAIRE
      fbBudSec1..fbBudSec5  : Err:='#ERREUR:Non implémenté en CWAS'; // A FAIRE
      {$ENDIF}
    {$ELSE}
      {$IFNDEF CCS3}
      fbBudGen : FicheBudGene(Nil,'',Qui,taConsult,0);
      fbBudSec1..fbBudSec5  : FicheBudSect(Nil,Axe,Qui,taConsult,0);
      {$ENDIF}
    {$ENDIF}
   fbJal  : FicheJournal(nil, '', Qui,taConsult, 0);
   fbRubrique : FicheRubrique(Nil,'',Qui,taConsult,0);
  END;
  FinZoomOLE ;
END ;
Result:=Err ;
END ;

{$IFDEF AMORTISSEMENT}
Function THalleyWindow.GetBaseTaxePro(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
// Nature bien = GEB, IAP, UNA
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result := AmGetBaseTaxePro(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup);
end ;

Function THalleyWindow.GetBaseTaxePro2(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup,Eligible : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
// Nature bien = GEB, IAP, UNA
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
Result := AmGetBaseTaxePro(CpteInf,CpteSup,NatureBien,LieuInf,LieuSup,Eligible);
end ;

Function THalleyWindow.GetAugAcqMut(CpteInf,CpteSup : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  Result:='#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;
  Result:=AmGetAugAcqMut(CpteInf,CpteSup);
end ;

Function THalleyWindow.GetDiminutionSortie(CpteInf,CpteSup : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;
  Result:=AmGetDiminutionSortie(CpteInf,CpteSup);
end ;

Function THalleyWindow.GetDiminutionMutation(CpteInf,CpteSup : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;
  Result:=AmGetDiminutionMutation(CpteInf,CpteSup);
end ;

Function THalleyWindow.GetDotationExo(CpteInf,CpteSup, Methode : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;
  Result:=AmGetDotationExo(CpteInf,CpteSup, Methode);
end ;

Function THalleyWindow.GetDiminutionAmort(CpteInf,CpteSup : string) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;
  Result:=AmGetDiminutionAmort(CpteInf,CpteSup);
end ;

Function THalleyWindow.GetListeLieuImmo : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;
  Result := AmGetListeLieuImmo;
end ;

Function THalleyWindow.GetValHTImmo( pstCpteInf, pstCpteSup, pstParam : string ) : Variant ;
begin
  Result := JaiLeDroit;
  if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  if Not V_PGI.OKOuvert then Exit ;
  Result := AmGetValHTImmo( pstCpteInf, pstCpteSup, pstParam);
end;
(*
Function THalleyWindow.Balance(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
Var Q : TQuery ;
    Sql,St,St1,St2 : String ;
    Err,Axe : String ;
    OkLib : Boolean ;
    fb : TFichierBase ;
    ii : Integer ;
    hh : hwnd ;
BEGIN
if(Table='') then BEGIN Result:='#ERREUR:Table non renseignée' ; Exit ; END ;
if(Qui='') then BEGIN Result:='#ERREUR:Code recherché non renseigné' ; Exit ; END ;
hh:=GetForeGroundWindow ;
MakeWindowActive(FMenuPrinc.Handle) ;
Result:=Get_Balance(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif) ;
MakeWindowActive(hh) ;
END ;
*)

// stType = 'ENC' ==> liste des immos en cours
// stType = 'SOR' ==> liste des immos sorties dans l'année
procedure THalleyWindow.ChargeListeImmo(stType : string);
var Q : TQuery;
    OBImmo , OBAmor : TOB;
    PlanInfo : TPlanInfo;
    PMV : TImPMValue;
    St : string;
    i : integer;
    a,m,j : word;
begin
  FListeImmo.Clear ; FPosListeImmo:=0 ;
  if Not V_PGI.OKOuvert then Exit ;
  if (stType = 'ENC') then
  begin
    OBImmo := TOB.Create ('', nil, - 1);
    PlanInfo := TPlanInfo.Create('');
    Q := OpenSQL ('SELECT I_PLANACTIF,I_IMMO,I_LIBELLE,I_COMPTEIMMO,I_DUREEREPRISE,'+
      'I_DATEPIECEA,I_DATEAMORT,I_BASEECO,I_BASEFISC,I_BASEAMORDEBEXO,I_BASEAMORFINEXO,'+
      'I_DATECREATION,I_MONTANTEXC,I_TYPEEXC,I_REINTEGRATION,I_QUOTEPART,I_TVARECUPERABLE, I_TVARECUPEREE, '+
      'I_VALEURACHAT,I_ETAT,I_MONTANTHT,I_REINTEGRATION,I_OPECHANGEPLAN,I_DATECESSION,'+
      'I_REVISIONECO,I_REPRISEECO,I_REPCEDECO,I_BASEECO,I_DUREEECO,I_METHODEECO,'+
      'I_TAUXECO,I_METHODEFISC,I_REVISIONFISCALE,I_REPRISEFISCAL,I_REPCEDFISC,'+
      'I_BASEFISC,I_DUREEFISC,I_METHODEFISC,I_TAUXFISC,I_OPECESSION FROM IMMO WHERE I_ETAT<>"FER" AND I_QUALIFIMMO="R" AND I_DATEPIECEA<="'+USDatetime(VH^.Encours.Fin)+'"',True);
    OBImmo.LoadDetailDB('IMMO','','',Q,False,True) ;
    Ferme (Q);
    for i:=0 to OBImmo.Detail.Count - 1 do
    begin
      // Chargement du plan d'amortissement
      Q := OpenSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+
         OBImmo.Detail[i].GetValue('I_IMMO')+'" ORDER BY IA_NUMEROSEQ DESC, IA_DATE',True);
      OBAmor := TOB.Create('',OBImmo.Detail[i],-1);
      OBAmor.LoadDetailDB('IMMOAMOR','','',Q,False,False);
      Ferme (Q);
      // Calcul des informations
//      PlanInfo.UpdateTOB (OBImmo.Detail[i].GetValue('I_IMMO'));
      PlanInfo.ChargeImmo(OBImmo.Detail[i].GetValue('I_IMMO'));
      PlanInfo.Calcul (VH^.Encours.Fin,false);
      // Mise à jour de la liste                             imedca
      DecodeDate (OBImmo.Detail[i].GetValue('I_DATEPIECEA'),a, m, j);
//     St := Format('%d'#9'%25.25s'#9'%.02d%.02d%.04d'#9'%.2f'#9'%.2f'#9'%.2f'#9'%.2f'#9'%.2f'#9'%.2f',
      St := Format('%d'#9'%25.25s'#9'%.02d%.02d%.04d'#9'%s'#9'%s'#9'%s'#9'%s'#9'%s'#9'%s',
        [FListeImmo.Count+1,VirePV(TrimLeft(OBImmo.Detail[i].GetValue('I_LIBELLE'))),j,m,a,
        StrFPoint(OBImmo.Detail[i].GetValue('I_MONTANTHT')),
        StrFPoint(OBImmo.Detail[i].GetValue('I_TVARECUPERABLE')),
        StrFPoint(OBImmo.Detail[i].GetValue('I_TVARECUPEREE')),
        StrFPoint(OBImmo.Detail[i].GetValue('I_TAUXECO')),
        StrFPoint(PlanInfo.CumulAntEco),StrFPoint(PlanInfo.DotationEco)]);
        FListeImmo.Add (St);
    end;
    PlanInfo.Free;
    OBImmo.Free;
  end else
  if (stType = 'SOR') then
  begin
    {Q := OpenSQL ('SELECT I_LIBELLE,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT,'+
      'I_METHODEECO,IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES'+
      ' FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE IL_TYPEOP LIKE "CE%" '+
      ' AND IL_DATEOP>="'+USDateTime(VH^.Encours.Deb)+'"'+
      ' AND IL_DATEOP<="'+USDatetime(VH^.Encours.Fin)+'"',True); }
    // BTY 04/06 FQ 17629 Récup colonne règle de calcul de la PMValue
    Q := OpenSQL ('SELECT I_LIBELLE,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT,'+
      'I_METHODEECO,I_REGLECESSION,IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES'+
      ' FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE IL_TYPEOP LIKE "CE%" '+
      ' AND IL_DATEOP>="'+USDateTime(VH^.Encours.Deb)+'"'+
      ' AND IL_DATEOP<="'+USDatetime(VH^.Encours.Fin)+'"',True);
    while not Q.Eof do
    begin
      DecodeDate (Q.FindField('IL_DATEOP').AsDateTime,a,m,j);
      {PMV := CalculPMValue (Q.FindField('IL_DATEOP').AsDateTime,
          Q.FindField('I_DATEPIECEA').AsDateTime,Q.FindField('I_DATEAMORT').AsDateTime,
          Q.FindField('IL_PVALUE').AsFloat,Q.FindField('IL_CUMANTCESECO').AsFloat,
          Q.FindField('IL_DOTCESSECO').AsFloat,Q.FindField('I_METHODEECO').AsString); }
      // BTY 04/06 FQ 17629 Ajout paramètre de calcul de la PMValue
      PMV := CalculPMValue (Q.FindField('IL_DATEOP').AsDateTime,
          Q.FindField('I_DATEPIECEA').AsDateTime,Q.FindField('I_DATEAMORT').AsDateTime,
          Q.FindField('IL_PVALUE').AsFloat,Q.FindField('IL_CUMANTCESECO').AsFloat,
          Q.FindField('IL_DOTCESSECO').AsFloat,Q.FindField('I_METHODEECO').AsString,
          Q.FindField('I_REGLECESSION').AsString );

//      St := Format ('%d'#9'%s'#9'%.02d%.02d%.04d'#9'%.2f'#9'%.2f'#9'%.2f',
      St := Format ('%d'#9'%s'#9'%.02d%.02d%.04d'#9'%s'#9'%s'#9'%s',
        [FListeImmo.Count+1,VirePV(TrimLeft(Q.FindField('I_LIBELLE').AsString)),j,m,a,StrFPoint(Q.FindField('IL_MONTANTCES').AsFloat),
        StrFPoint(PMV.PCT-PMV.MCT),StrFPoint(PMV.PLT-PMV.MLT)]);
        FListeImmo.Add (St);
      Q.Next;
    end;
    Ferme (Q);
  end;
end;

function THalleyWindow.GetNextImmo: Variant;
var stNext : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  stNext:='' ;
  if Not V_PGI.OKOuvert then Exit ;
  if FPosListeImmo>=FListeImmo.Count then exit ;
  stNext:=FListeImmo[FPosListeImmo] ;
  Inc(FPosListeImmo) ;
  result := stNext;
end;

// stType = 'ENC' ==> liste des immos en cours
// stType = 'SOR' ==> liste des immos sorties dans l'année
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/01/2003
Modifié le ... :   /  /
Description .. : Nouvelle procédure de chargement de la liste des
Suite ........ : immobilisations.
Suite ........ : Mise en place d'un numéro de version
Mots clefs ... :
*****************************************************************}
function THalleyWindow.ChargeLesImmos(stType : string; iVersion : integer) : variant;
var Q : TQuery;
    OBImmo , OBAmor : TOB;
    PlanInfo : TPlanInfo;
    PMV : TImPMValue;
    St : string;
    i : integer;
    a,m,j,a1,m1,j1 : word;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  FListeImmo.Clear ; FPosListeImmo:=0 ;
  if Not V_PGI.OKOuvert then Exit ;
  if (stType = 'ENC') then
  begin
    OBImmo := TOB.Create ('', nil, - 1);
    PlanInfo := TPlanInfo.Create('');
    Q := OpenSQL ('SELECT I_PLANACTIF,I_IMMO,I_LIBELLE,I_COMPTEIMMO,I_DUREEREPRISE,'+
      'I_DATEPIECEA,I_DATEAMORT,I_BASEECO,I_BASEFISC,I_BASEAMORDEBEXO,I_BASEAMORFINEXO,'+
      'I_DATECREATION,I_MONTANTEXC,I_TYPEEXC,I_REINTEGRATION,I_QUOTEPART,I_TVARECUPERABLE, I_TVARECUPEREE, '+
      'I_VALEURACHAT,I_ETAT,I_MONTANTHT,I_REINTEGRATION,I_OPECHANGEPLAN,I_DATECESSION,'+
      'I_REVISIONECO,I_REPRISEECO,I_REPCEDECO,I_BASEECO,I_DUREEECO,I_METHODEECO,'+
      'I_TAUXECO,I_METHODEFISC,I_REVISIONFISCALE,I_REPRISEFISCAL,I_REPCEDFISC,'+
      'I_BASEFISC,I_DUREEFISC,I_METHODEFISC,I_TAUXFISC,I_OPECESSION FROM IMMO WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_ETAT<>"FER" AND I_QUALIFIMMO="R" AND I_DATEPIECEA<="'+USDatetime(VH^.Encours.Fin)+'"',True);
    OBImmo.LoadDetailDB('IMMO','','',Q,False,True) ;
    Ferme (Q);
    for i:=0 to OBImmo.Detail.Count - 1 do
    begin
      // Chargement du plan d'amortissement
      Q := OpenSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+
         OBImmo.Detail[i].GetValue('I_IMMO')+'" ORDER BY IA_NUMEROSEQ DESC, IA_DATE',True);
      OBAmor := TOB.Create('',OBImmo.Detail[i],-1);
      OBAmor.LoadDetailDB('IMMOAMOR','','',Q,False,False);
      Ferme (Q);
      // Calcul des informations
//      PlanInfo.UpdateTOB (OBImmo.Detail[i].GetValue('I_IMMO'));
      PlanInfo.ChargeImmo(OBImmo.Detail[i].GetValue('I_IMMO'));
      PlanInfo.Calcul (VH^.Encours.Fin,false);
      // Mise à jour de la liste
      DecodeDate (OBImmo.Detail[i].GetValue('I_DATEPIECEA'),a, m, j);
//     St := Format('%d'#9'%25.25s'#9'%.02d%.02d%.04d'#9'%.2f'#9'%.2f'#9'%.2f'#9'%.2f'#9'%.2f'#9'%.2f%1s',
      St := Format('%d'#9'%s'#9'%.02d%.02d%.04d'#9'%s'#9'%s'#9'%s'#9'%s'#9'%s'#9'%s'#9'%s',
        [FListeImmo.Count+1,VirePV(TrimLeft(OBImmo.Detail[i].GetValue('I_LIBELLE'))),j,m,a,
        StrFPoint(OBImmo.Detail[i].GetValue('I_MONTANTHT')),
        StrFPoint(OBImmo.Detail[i].GetValue('I_TVARECUPERABLE')),
        StrFPoint(OBImmo.Detail[i].GetValue('I_TVARECUPEREE')),
        Copy(OBImmo.Detail[i].GetValue('I_METHODEECO'),1,1),
        StrFPoint(OBImmo.Detail[i].GetValue('I_TAUXECO')),
        StrFPoint(PlanInfo.CumulAntEco),StrFPoint(PlanInfo.DotationEco)]);
        FListeImmo.Add (St);
    end;
    PlanInfo.Free;
    OBImmo.Free;
  end else
  if (stType = 'SOR') then
  begin
    {Q := OpenSQL ('SELECT I_LIBELLE,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT,'+
      'I_METHODEECO,IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES,IL_VOCEDEE'+
      ' FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_QUALIFIMMO="R" AND IL_TYPEOP LIKE "CE%" '+
      ' AND IL_DATEOP>="'+USDateTime(VH^.Encours.Deb)+'"'+
      ' AND IL_DATEOP<="'+USDatetime(VH^.Encours.Fin)+'"',True); }
   // BTY 04/06 FQ 17629 Récup colonne règle de calcul de la PMValue
    Q := OpenSQL ('SELECT I_LIBELLE,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT,'+
      'I_METHODEECO,I_REGLECESSION,IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES,IL_VOCEDEE'+
      ' FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE (I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_QUALIFIMMO="R" AND IL_TYPEOP LIKE "CE%" '+
      ' AND IL_DATEOP>="'+USDateTime(VH^.Encours.Deb)+'"'+
      ' AND IL_DATEOP<="'+USDatetime(VH^.Encours.Fin)+'"',True);

    while not Q.Eof do
    begin
//      DecodeDate (Q.FindField('IL_DATEOP').AsDateTime,a,m,j);
      DecodeDate (Q.FindField('I_DATEPIECEA').AsDateTime,a,m,j);
//      DecodeDate (Q.FindField('I_DATEAMORT').AsDateTime,a1,m1,j1);
      { CA - 20/01/2004 - Suite demande JPO pour liasses fiscales }
      DecodeDate (Q.FindField('IL_DATEOP').AsDateTime,a1,m1,j1);
      {PMV := CalculPMValue (Q.FindField('IL_DATEOP').AsDateTime,
          Q.FindField('I_DATEPIECEA').AsDateTime,Q.FindField('I_DATEAMORT').AsDateTime,
          Q.FindField('IL_PVALUE').AsFloat,Q.FindField('IL_CUMANTCESECO').AsFloat,
          Q.FindField('IL_DOTCESSECO').AsFloat,Q.FindField('I_METHODEECO').AsString); }
      // BTY 04/06 FQ 17629 Ajout paramètre de calcul de la PMValue
      PMV := CalculPMValue (Q.FindField('IL_DATEOP').AsDateTime,
             Q.FindField('I_DATEPIECEA').AsDateTime,Q.FindField('I_DATEAMORT').AsDateTime,
             Q.FindField('IL_PVALUE').AsFloat,Q.FindField('IL_CUMANTCESECO').AsFloat,
             Q.FindField('IL_DOTCESSECO').AsFloat,Q.FindField('I_METHODEECO').AsString,
             Q.FindField('I_REGLECESSION').AsString  );

//      St := Format ('%d'#9'%s'#9'%.02d%.02d%.04d'#9'%.2f'#9'%.2f'#9'%.2f',
      St := Format ('%d'#9'%s'#9'%.02d%.02d%.04d'#9'%.02d%.02d%.04d'#9'%s'#9'%s'#9'%s'#9'%s'#9'%s',
        [FListeImmo.Count+1,   // indice immobilisation
        VirePV(TrimLeft(Q.FindField('I_LIBELLE').AsString)),  // Libellé
        j1,m1,a1,  // Date d'opération
        j,m,a,  // Date d'achat
        StrFPoint(Q.FindField('IL_VOCEDEE').AsFloat),       // Valeur d'origine
        StrFPoint(Q.FindField('IL_CUMANTCESECO').AsFloat+Q.FindField('IL_DOTCESSECO').AsFloat),// Cumul des dotations
        StrFPoint(Q.FindField('IL_MONTANTCES').AsFloat),  // Montant de la sortie
        StrFPoint(PMV.PCT-PMV.MCT),     // +Value à court-terme
        StrFPoint(PMV.PLT-PMV.MLT)]);   // +Value à long-terme
        FListeImmo.Add (St);
      Q.Next;
    end;
    Ferme (Q);
  end;
  Result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/01/2003
Modifié le ... :   /  /
Description .. : Nouvelle procédure de récupération d'une immobilisation
Suite ........ : avec gestion du numéro de version
Mots clefs ... :
*****************************************************************}
function THalleyWindow.ImmoSuivante ( iVersion : integer ) : Variant;
var stNext : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  stNext:='' ;
  if Not V_PGI.OKOuvert then Exit ;
  if FPosListeImmo>=FListeImmo.Count then exit ;
  stNext:=FListeImmo[FPosListeImmo] ;
  Inc(FPosListeImmo) ;
  result := stNext;
end;

{$ENDIF}

function THalleyWindow.GetUserName : Variant;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Non connecté' ;
if Not V_PGI.OKOuvert then Exit ;
result := V_PGI.UserLogin;
end;

{========================================================================================}
{========================== EXCEL TO PGI ================================================}
{========================================================================================}
Function THalleyWindow.TraiteEcrituresRevision ( LeJournal,LeFolio : string ) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=IntegreEcritureRevisionPGI(LeJournal,LeFolio) ;
end;

// CA - 29/11/2001 - Génère écriture avec création des comptes
Function THalleyWindow.TraiteEcrituresComptesRevision ( LeJournal,LeFolio : string ) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=IntegreEcritureRevisionPGI(LeJournal,LeFolio,True) ;
end;

// ajout me 21/06/2002
Function THalleyWindow.TraiteEcrituresComptesRevisionach ( LeJournal,LeFolio : string ) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=IntegreEcritureRevisionPGI(LeJournal,LeFolio,True,'ACH') ;
end;

Function THalleyWindow.TraiteEcrituresComptesRevisionvte ( LeJournal,LeFolio : string ) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=IntegreEcritureRevisionPGI(LeJournal,LeFolio,True,'VTE') ;
end;

Function THalleyWindow.TraiteEcrituresComptesRevisionbqe ( LeJournal,LeFolio,Contrepartie : string ) : Variant ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=IntegreEcritureRevisionPGI(LeJournal,LeFolio,True,'BQE',Contrepartie) ;
end;

// ajout me 16-10-2002
Function THalleyWindow.TraiteEcrituresBalance ( LeJournal,LeFolio,RecupLib : string) : Variant ;
var
Rec      : Boolean;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Rec := (RecupLib = 'TRUE');
Result:=IntegreEcritureRevisionPGI(LeJournal,LeFolio,True, '', '', Rec) ;
end;

Function THalleyWindow.GetListeCollectif (Nature :string) : Variant ;
var s : string ;
    Q : TQuery ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Liste collectif non trouvée' ;
if Not V_PGI.OKOuvert then Exit ;
s:='' ;
Q:=OpenSQL('SELECT G_GENERAL from GENERAUX Where G_NATUREGENE="'+ Nature +'"',TRUE) ;
While Not Q.Eof Do
  BEGIN
  s:=s+Q.FindField('G_GENERAL').AsString+';' ;
  Q.Next ;
  END ;
Ferme(Q) ;
Result:=s ;
VH^.CPAppelOLE:=True ;
end ;

Function THalleyWindow.GetListeTiers (Coll :string) : Variant ;
var s    : string ;
    Q    : TQuery ;
    gene : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
result := '#Erreur : Liste tiers non trouvée' ;
if Not V_PGI.OKOuvert then Exit ;
s:='' ;
Gene:=BourreEtLess(Coll,fbGene);
Q:=OpenSQL('SELECT T_AUXILIAIRE,T_COLLECTIF from TIERS Where T_COLLECTIF="'+ Gene +'"',TRUE) ;
While Not Q.Eof Do
  BEGIN
  s:=s+Q.FindField('T_AUXILIAIRE').AsString+';' ;
  Q.Next ;
  END ;
Ferme(Q) ;
Result:=s ;
VH^.CPAppelOLE:=True ;
end ;


function THalleyWindow.ChargeJournaux : variant;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=ChargeJournauxRevisionPGI;
end;

// ajout me 16/06/20004
function THalleyWindow.ChargeSection ( Ax : String ) : variant;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=ChargeSectionsRevisionPGI (Ax);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : M. ENTRESSANGLE
Créé le ...... : 16/06/2004
Modifié le ... :   /  /
Description .. : Chargement des ventilation type
Mots clefs ... :
*****************************************************************}
function THalleyWindow.ChargeVentilType : variant;
var V : Variant;
    It, Va : TStrings;
    i : integer;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=varEmpty ;
It:=TStringList.Create ; Va:=TStringList.Create ;
RemplirValCombo('TTVENTILTYPE','','',It,Va,False,False);
V:=VarArrayCreate([0,It.Count],varVariant);
for i:=0 to It.Count-1 do  V[i]:=Va[i];
//V[i]:=Va[i]+' : '+It[i] ;
Result:=V ;
It.Free ; Va.Free ;
END ;

function THalleyWindow.ChargeEtablissement : variant;
var V : Variant;
    It, Va : TStrings;
    i : integer;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=varEmpty ;
It:=TStringList.Create ; Va:=TStringList.Create ;
RemplirValCombo('TTETABLISSEMENT','','',It,Va,False,False);
V:=VarArrayCreate([0,It.Count],varVariant);
for i:=0 to It.Count-1 do  V[i]:=Va[i];
Result:=V ;
It.Free ; Va.Free ;
END ;

function THalleyWindow.GetFolios( stJournal: string ): Variant;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
Result:=GetFoliosRevisionPGI(stJournal);
end;

function THalleyWindow.GetSeriaRevision: Variant;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  if VH^.OkModRevision then Result := 'X'
  else Result := '-';
end;

function THalleyWindow.GetBalanceSituation(TypeCum, CodeBal: string): Variant;
var Q : TQuery;
    St : string;
    stLibelle, stJointure : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  // Récupère la liste des balances de situation
  if Not V_PGI.OKOuvert then Exit ;
  if TypeCum = '' then Exit;
  St := '' ;
  // Récupère les écritures de la balance de situation
  if (TypeCum = 'GEN') or (TypeCum='G/T') then
  begin
    stLibelle := ',G_LIBELLE';
    stJointure := ' LEFT JOIN GENERAUX ON G_GENERAL=BSE_COMPTE1';
  end else if TypeCum = 'TIE' then
  begin
    stLibelle := ',T_LIBELLE';
    stJointure := ' LEFT JOIN TIERS ON T_AUXILIAIRE=BSE_COMPTE1';
  end else if (TypeCum = 'ANA') or (TypeCum='G/A') then
  begin
    stLibelle := ',S_LIBELLE';
    stJointure := ' LEFT JOIN SECTION ON S_SECTION=BSE_COMPTE1';
  end;
  if (stLibelle='') or (stJointure='') then exit;
  Q := OpenSQL ('SELECT CBALSITECR.*'+stLibelle+' LIBELLE FROM CBALSITECR '+stJointure+
        ' WHERE BSE_CODEBAL="'+CodeBal+'" ORDER BY BSE_COMPTE1,BSE_COMPTE2',True);
  while not Q.Eof do
  begin
    St := St + Trim(Q.FindField('BSE_COMPTE1').AsString)+ ';' ;
    // Si compte2 vide, on ajoute un espace pour JP
    if Q.FindField('BSE_COMPTE2').AsString='' then St := St+' ;'
    else St := St + Trim(Q.FindField('BSE_COMPTE2').AsString) + ';';
    St := St + VirePV(Trim(Q.FindField('LIBELLE').AsString))+ ';';
    St := St + Strfpoint(Q.FindField('BSE_DEBIT').AsFloat) + ';';
    St := St + Strfpoint(Q.FindField('BSE_CREDIT').AsFloat);
    St := St + #10;
    Q.Next;
  end;
  Ferme (Q);
  Result := St;
  VH^.CPAppelOLE:=True ;
end;

function THalleyWindow.GetListeSituation ( TypeCum, Exo : string ) : Variant;
var St , stWhere: string;
    Q : TQuery;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  result := '#Erreur : Non connecté' ;
  // Récupère la liste des balances de situation
  if Not V_PGI.OKOuvert then Exit ;
  if TypeCum = '' then Exit;
  St := '' ; stWhere := '';
  if TypeCum <> '' then stWhere := ' WHERE BSI_TYPECUM="'+TypeCum+'"';
  if Exo<>'' then
  begin
    if stWhere <> '' then stWhere := stWhere + ' AND BSI_EXERCICE="'+Exo+'" OR BSI_EXERCICE="---"'
    else stWhere := ' WHERE BSI_EXERCICE="'+Exo+'" OR BSI_EXERCICE="---"';
  end;
  Q := OpenSQL ('SELECT BSI_CODEBAL,BSI_LIBELLE,BSI_EXERCICE,BSI_DATE1,BSI_DATE2 FROM CBALSIT' + stWhere, True);
  while not Q.Eof do
  begin
    St := St + Q.FindField('BSI_CODEBAL').AsString+ ';' ;
    St := St + VirePV(Q.FindField('BSI_LIBELLE').AsString)+ ';' ;
    St := St + FormatDateTime('ddmmyyyy',(Q.FindField('BSI_DATE1').AsDateTime)) + ';';
    St := St + FormatDateTime('ddmmyyyy',(Q.FindField('BSI_DATE2').AsDateTime));
    St := St + #10;
    Q.Next;
  end;
  Ferme (Q);
  Result := St;
  VH^.CPAppelOLE:=True ;
end;

function THalleyWindow.GetResultatSituation ( CodeBal : string ) : Variant;
var SQL : string;
    Q   : TQuery;
    XPro, XCha : double;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
  XPro := 0; XCha := 0;
  result := '#Erreur : Non connecté' ;
  // Produits
  if Not V_PGI.OKOuvert then Exit ;
  SQL:='SELECT SUM(BSE_CREDIT-BSE_DEBIT) FROM CBALSITECR LEFT JOIN GENERAUX ON BSE_COMPTE1=G_GENERAL '
      +'WHERE BSE_CODEBAL="'+CodeBal+'" AND G_NATUREGENE="PRO"' ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then XPro:=Q.Fields[0].AsFloat ;
  Ferme(Q) ;
  // Charges
  SQL:='SELECT SUM(BSE_DEBIT-BSE_CREDIT) FROM CBALSITECR LEFT JOIN GENERAUX ON BSE_COMPTE1=G_GENERAL '
      +'WHERE BSE_CODEBAL="'+CodeBal+'" AND G_NATUREGENE="CHA"' ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then XCha:=Q.Fields[0].AsFloat ;
  Ferme(Q) ;
  Result:=Arrondi(XPro-XCha,V_PGI.OkDecV) ;
end;

function THalleyWindow.ComptaInitialisee: variant;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
{$IFNDEF EAGLCLIENT}
  Result := GetFlagAppli(ExtractFileName(Application.ExeName));
{$ELSE}
  result := False;
{$ENDIF}
end;

procedure THalleyWindow.MaximiseCompta;
begin
  Application.Restore;
end;

procedure THalleyWindow.MinimiseCompta;
begin
  Application.Minimize;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Djamel ZEDIAR
Créé le ...... : 18/04/2002
Modifié le ... :   /  /
Description .. : Verbe OLE Pour CRE
Mots clefs ... :
*****************************************************************}
Function THalleyWindow.GetInfoCRE : Variant;
var
   x:string;
   lObQuery       : TQuery;
   lInLastEmpCode : Integer;
   lDtDateEch     : TDateTime;
   lRdAmoEch      : Double;
   lRdIntEch      : Double;
   lRdAssEch      : Double;
   lRdIntEchPrec  : Double;
   lRdAssEchPrec  : Double;
   lDtDate1       : TDateTime;
   lDtDate2       : TDateTime;
   lDtDate3       : TDateTime;
   lDtDate4       : TDateTime;
   lDtDateEchPrec : TDateTime;
   lRdProrata     : Double;
   lRdTotalExe    : Double;
   lRdTotalMoins1 : Double;
   lRdTotal1a5    : Double;
   lRdTotalPlus5  : Double;
   lRdIntExe      : Double;
   lRdIntApresExe : Double;
   lRdAssExe      : Double;
   lRdIntCourus1  : Double;
   lRdIntCourus2  : Double;
   lRdAssCourus1  : Double;
   lRdAssCourus2  : Double;
   lRdIntCCA1     : Double;
   lRdIntCCA2     : Double;
   lRdAssCCA1     : Double;
   lRdAssCCA2     : Double;
   lRdAssApresExe : Double;
   lInAn,
   lInMois ,
   lInJour,lInJour2 : Word;
   lBoRupture     : Boolean;
   lStSQL         : String;
   lStValeurs    : String;
   lStNomsChamps : String;
   lObResult     : TStringList;
   lBoManquant   : Boolean;

   //
   Procedure Stocker(pStNomChamps : String; pStValeur : String);
   begin
      pStNomChamps := Trim(Uppercase(pStNomChamps));

      if lStNomsChamps <> '' then
      begin
         lStNomsChamps := lStNomsChamps + cSeparateur;
         lStValeurs := lStValeurs + cSeparateur;
      end
      else
         if lStValeurs <> '' then
            lStValeurs := lStValeurs + cRetourChariot;

      lStNomsChamps := lStNomsChamps + pStNomChamps;
      lStValeurs := lStValeurs + pStValeur;
   end;

begin
Result := JaiLeDroit;
if Result <> '' then exit;
   // LG - 23/09/2004
   lBoManquant   := False;
   lRdIntEchPrec := 0;
   lRdAssEchPrec := 0;

   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
   lObResult := Nil;
   lStValeurs := '';

   //
   // Calcul des dates
   //
   // Date 1 : début d'exercice
    lDtDate1 := VH^.Encours.Deb;
   decodedate(lDtDate1, lInAn, lInMois, lInJour);
   // Date 2 : début d'exercice suivant
   try
    //Djamel : date début exercice N+1 = date début exercice N + 12 mois
    //  lDtDate2 := EncodeDate(lInAn + 1, lInMois, lInJour);
    (*CAD : date début exercice N+1 = date fin exercice N + 1 jour*)
    lDtDate2 := VH^.Encours.Fin;
    lDtDate2 := lDtDate2 + 1;
    decodedate(lDtDate2, lInAn, lInMois, lInJour2);
   except
      lDtDate2 := EncodeDate(lInAn + 1, lInMois, 28);
   end;
   // Date 3 : Exercice suivant
   try
   //   lDtDate3 := EncodeDate(lInAn + 2, lInMois, lInJour);
      lDtDate3 := EncodeDate(lInAn + 1, lInMois, lInJour);
   except
   //   lDtDate3 := EncodeDate(lInAn + 2, lInMois, 28);
      lDtDate3 := EncodeDate(lInAn + 1, lInMois, 28);
   end;
   // Date 4 : + d'un an et - de 5 ans
   try
  //    lDtDate4 := EncodeDate(lInAn + 6, lInMois, lInJour);
      lDtDate4 := EncodeDate(lInAn + 5, lInMois, lInJour);
   except
  //    lDtDate4 := EncodeDate(lInAn + 6, lInMois, 28);
      lDtDate4 := EncodeDate(lInAn + 5, lInMois, 28);
   end;

   try
      lStSQL := ' Select ECH_DATE, ECH_AMORTISSEMENT, ECH_INTERET,' +
                '  ECH_ASSURANCE, ECH_VERSEMENT, ECH_DETTE,' +
                '  ECH_DATEEXERCICE, EMP_CODEEMPRUNT, EMP_LIBEMPRUNT,' +
                '  EMP_NUMCOMPTE, EMP_DATEDEBUT, EMP_DATECONTRAT,' +
                '  EMP_CAPITAL, EMP_DUREE, EMP_TOTINT,' +
                '  EMP_TOTASS, EMP_TOTAMORT, EMP_TOTVERS,' +
                '  EMP_NBECHTOT, G_GENERAL, G_LIBELLE' +
                ' From FECHEANCE, FEMPRUNT' +
                ' Left Join GENERAUX On G_GENERAL = EMP_NUMCOMPTE' +
                ' Where ECH_CODEEMPRUNT = EMP_CODEEMPRUNT' +
                  ' And EMP_CODEEMPRUNT > -1' +
                ' Order By EMP_CODEEMPRUNT, ECH_DATE';

      lObQuery := OpenSQL(lStSQL, True);

      lRdTotalExe    := 0;
      lRdTotalMoins1 := 0;
      lRdTotal1a5    := 0;
      lRdTotalPlus5  := 0;

      lRdIntExe      := 0;
      lRdIntCourus1  := 0;
      lRdIntCourus2  := 0;
      lRdIntApresExe := 0;

      lRdAssExe      := 0;
      lRdAssCourus1  := 0;
      lRdAssCourus2  := 0;
      lRdAssApresExe := 0;

      lRdIntCCA1  := 0;
      lRdIntCCA2  := 0;
      lRdAssCCA1  := 0;
      lRdAssCCA2  := 0;

      lDtDateEchPrec := -1;

      lInLastEmpCode := -1000;

      //
      //  Boucle
      //
      While not lObQuery.Eof do
      begin
         // Entête de bande détail
         {$IFNDEF EAGLCLIENT}
         if lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger <> lInLastEmpCode then
         {$ELSE}
         if lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger <> lInLastEmpCode then
         {$ENDIF}
         begin
            lStNomsChamps := '';

            lRdTotalExe    := 0;
            lRdTotalMoins1 := 0;
            lRdTotal1a5    := 0;
            lRdTotalPlus5  := 0;

            lRdIntExe      := 0;
            lRdIntCourus1  := 0;
            lRdIntCourus2  := 0;
            lRdIntApresExe := 0;

            lRdAssExe      := 0;
            lRdAssCourus1  := 0;
            lRdAssCourus2  := 0;
            lRdAssApresExe := 0;

            lRdIntCCA1 := 0;
            lRdIntCCA2 := 0;
            lRdAssCCA1 := 0;
            lRdAssCCA2 := 0;


         //   lDtDateEchPrec := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime;
//Correction CAD du 07/12/04 : la date de la première échéance ne doit pas être celle de la date du contrat !!!
         {$IFNDEF EAGLCLIENT}
            lDtDateEchPrec := -1;
            lBoManquant := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime < lObQuery.FindField('EMP_DATEDEBUT').AsDateTime;

            Stocker('NumEmprunt',  lObQuery.FindField('EMP_CODEEMPRUNT').AsString);
            Stocker('LibEmp',      lObQuery.FindField('EMP_LIBEMPRUNT').AsString);
            Stocker('NumCompte',   lObQuery.FindField('EMP_NUMCOMPTE').AsString);
            Stocker('LibCompte',   lObQuery.FindField('G_LIBELLE').AsString);
            Stocker('DateContrat', lObQuery.FindField('EMP_DATECONTRAT').AsString);
            Stocker('Duree',       lObQuery.FindField('EMP_DUREE').AsString);
            Stocker('Capital',     lObQuery.FindField('EMP_CAPITAL').AsString);
            if (lObQuery.FindField('EMP_DATECONTRAT').AsDateTime >= lDtDate1) and
               (lObQuery.FindField('EMP_DATECONTRAT').AsDateTime <  lDtDate2) then
               Stocker('CapNouvEmp', lObQuery.FindField('EMP_CAPITAL').AsString)
            else
               Stocker('CapNouvEmp', '0');
         end;

         lDtDateEch := lObQuery.FindField('ECH_DATE').AsDateTime;
         lRdAmoEch  := lObQuery.FindField('ECH_AMORTISSEMENT').AsFloat;
         lRdIntEch  := lObQuery.FindField('ECH_INTERET').AsFloat;
         lRdAssEch  := lObQuery.FindField('ECH_ASSURANCE').AsFloat;
         {$ELSE}
            lDtDateEchPrec := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime;
            lBoManquant := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime < lObQuery.FindField('EMP_DATEDEBUT').AsDateTime;
            Stocker('NumEmprunt', lObQuery.FindField('EMP_CODEEMPRUNT').AsString);
            Stocker('LibEmp', lObQuery.FindField('EMP_LIBEMPRUNT').AsString);
            Stocker('NumCompte', lObQuery.FindField('EMP_NUMCOMPTE').AsString);
            Stocker('LibCompte', lObQuery.FindField('G_LIBELLE').AsString);
            Stocker('DateContrat', lObQuery.FindField('EMP_DATECONTRAT').AsString);
            Stocker('Duree', lObQuery.FindField('EMP_DUREE').AsString);
            Stocker('Capital', lObQuery.FindField('EMP_CAPITAL').AsString);
            if (lObQuery.FindField('EMP_DATECONTRAT').AsDateTime >= lDtDate1) and
               (lObQuery.FindField('EMP_DATECONTRAT').AsDateTime < lDtDate2) then
               Stocker('CapNouvEmp', lObQuery.FindField('EMP_CAPITAL').AsString)
            else
               Stocker('CapNouvEmp', '0');
         end;
         lDtDateEch := lObQuery.FindField('ECH_DATE').AsDateTime;
         lRdAmoEch  := lObQuery.FindField('ECH_AMORTISSEMENT').AsFloat;
         lRdIntEch  := lObQuery.FindField('ECH_INTERET').AsFloat;
         lRdAssEch  := lObQuery.findField('ECH_ASSURANCE').AsFloat;
         {$ENDIF}

         // Calcul des totaux
         if (lDtDateEch >= lDtDate1) and (lDtDateEch < lDtDate2) then
         begin
            lRdTotalExe := lRdTotalExe + lRdAmoEch;
            lRdIntExe   := lRdIntExe + lRdIntEch;
            lRdAssExe := lRdAssExe + lRdAssEch;
         end
         else
            if (lDtDateEch >= lDtDate2) and (lDtDateEch < lDtDate3) then
            begin
               lRdTotalMoins1 := lRdTotalMoins1 + lRdAmoEch;
            end
            else
               if (lDtDateEch >= lDtDate3) and (lDtDateEch < lDtDate4) then
                  lRdTotal1a5 := lRdTotal1a5 + lRdAmoEch
               else
                  if lDtDateEch >= lDtDate4 then
                     lRdTotalPlus5 := lRdTotalPlus5 + lRdAmoEch;

         if lDtDateEch >= lDtDate2 then
         begin
            lRdIntApresExe := lRdIntApresExe + lRdIntEch;
            lRdAssApresExe := lRdAssApresExe + lRdAssEch;
         end;

         //
         // Courus N - 1
         //
         // si la date de fin d'exercice N-1 est entre l'échéance précédente
         // et l'échéance courante
         if lDtDateEchPrec>-1 then
         begin
             if (lDtDate1 - 1 > lDtDateEchPrec) and (lDtDate1 - 1 <= lDtDateEch) then
             begin
                if lBoManquant then
                begin
                   // Intérets manquants
                   lRdProrata := (lDtDate1 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCourus1 := lRdProrata * lRdIntEch;
                   lRdAssCourus1 := lRdProrata * lRdAssEch;
                end
                else
                begin
                   lRdProrata := (lDtDateEch - lDtDate1 + 1) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCCA1 := lRdProrata * lRdIntEchPrec;
                   lRdAssCCA1 := lRdProrata * lRdAssEchPrec;
                end;
             end;
              //
             // Courus N
             //
             // si la date de fin d'exercice est entre l'échéance précédente
             // et l'échéance courante
             if (lDtDate2 - 1 > lDtDateEchPrec) and (lDtDate2 - 1 <= lDtDateEch) then
             begin
                if lBoManquant then
                begin
                   // Intérets manquants
                   lRdProrata := (lDtDate2 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCourus2 := lRdProrata * lRdIntEch;
                   lRdAssCourus2 := lRdProrata * lRdAssEch;
                end
                else
                begin
                   lRdProrata := (lDtDateEch - lDtDate2 + 1) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCCA2 := lRdProrata * lRdIntEchPrec;
                   lRdAssCCA2 := lRdProrata * lRdAssEchPrec;
                end;
             end;
         end
      //******************************************************
         else
         begin
         //
         // Courus N
         //
         // 1. Si la date de mise à disposition des fonds est antérieure à la date de première échéance et
         // 2  si la date de mise à disposition des fonds se situe dans l'exercice et
         // 3  si la date de la première échéance est postérieure à l'exercice.
            if lBoManquant then
            begin
                {$IFNDEF EAGLCLIENT}
                lDtDateEchPrec := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime;
                {$ELSE}
                lDtDateEchPrec := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime;
                {$ENDIF}
             // si la date de fin d'exercice est entre l'échéance précédente
             // et l'échéance courante
                 if (lDtDate1 - 1 < lDtDateEchPrec) and (lDtDate2 - 1 > lDtDateEchPrec) and(lDtDate2 - 1 <= lDtDateEch) then
                 begin
                                // Intérets manquants
                   lRdProrata := (lDtDate2 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCourus2 := lRdProrata * lRdIntEch;
                   lRdAssCourus2 := lRdProrata * lRdAssEch;
                end;
            end;
         end;
      //******************************************************
         // Echéance suivante
         lDtDateEchPrec := lDtDateEch;
         lRdIntEchPrec  := lRdIntEch;
         lRdAssEchPrec  := lRdAssEch;
         {$IFNDEF EAGLCLIENT}
         lInLastEmpCode := lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger;
         {$ELSE}
          lInLastEmpCode := lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger;
          {$ENDIF}
         lObQuery.Next;

         // pied de bande détail
         if lObQuery.Eof then
            lBoRupture := True
         else
            {$IFNDEF EAGLCLIENT}
            lBoRupture := (lInLastEmpCode <> lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger);
            {$ELSE}
            lBoRupture := (lInLastEmpCode <> lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger);
            {$ENDIF}
         if lBoRupture then
         begin
            Stocker('AmoExe',   FloatTostr(lRdTotalExe));
            Stocker('Moins1an', FloatTostr(lRdTotalMoins1));
            Stocker('De1a5an',  FloatTostr(lRdTotal1a5));
            Stocker('Plus5ans', FloatTostr(lRdTotalPlus5));

            Stocker('IntExe',      FloatTostr(lRdIntExe));
            Stocker('IntCourus1',  FloatTostr(lRdIntCourus1));
            Stocker('IntCCA1',     FloatTostr(lRdIntCCA1));
            Stocker('IntCourus2',  FloatTostr(lRdIntCourus2));
            Stocker('IntCCA2',     FloatTostr(lRdIntCCA2));
            Stocker('IntApresExe', FloatTostr(lRdIntApresExe));

            Stocker('AssExe',      FloatTostr(lRdAssExe));
            Stocker('AssCourus1',  FloatTostr(lRdAssCourus1));
            Stocker('AssCCA1',     FloatTostr(lRdAssCCA1));
            Stocker('AssCourus2',  FloatTostr(lRdAssCourus2));
            Stocker('AssCCA2',     FloatTostr(lRdAssCCA2));
            Stocker('AssApresExe', FloatTostr(lRdAssApresExe));
         end;
      end;

      // Résultat
      Result := lStNomsChamps + cRetourChariot + lStValeurs;
      lObResult := TStringList.Create;
      try
         lObResult.Text := result;
         lObResult.SaveToFile('C:\CRE.CSV');
      except
      end;

   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
      if Assigned(lObResult) then lObResult.Free;;
   end;
end;

function THalleyWindow.JaiLeDroit: string;
begin
  if not ExJaiLeDroitConcept(ccOLE, FALSE) then
    Result := 'Vous n''avez pas les droits d''utilisation du lien OLE'
  else Result:= '';
  Result := Trim (Result);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 08/01/2003
Modifié le ... : 08/01/2003
Description .. : Mise à jour des soldes des comptes auxiliaires et généraux
Mots clefs ... : SOLDE;GENERAUX;COMPTE
*****************************************************************}
procedure THalleyWindow.MajSoldeComptes;
begin
  MajTotComptes([fbGene,fbAux],FALSE,FALSE,'') ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : M. ENTRESSANGLE
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. : Rend la nature du journal
Mots clefs ... :
*****************************************************************}
Function THalleyWindow.GetNatureJrl (Journal :string) : Variant ;
var s : string ;
    Q : TQuery ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
if Not V_PGI.OKOuvert then Exit ;
s:='' ;
Q:=OpenSQL('SELECT J_NATUREJAL from JOURNAL Where J_JOURNAL="'+ Journal +'"',TRUE) ;
if Not Q.Eof then
  BEGIN
  s:=Q.FindField('J_NATUREJAL').AsString;
  END ;
Ferme(Q) ;
Result:=s ;
VH^.CPAppelOLE:=True ;
end ;

Function THalleyWindow.GetContrepartieJrl (Journal :string) : Variant ;
var s : string ;
    Q : TQuery ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
if Not V_PGI.OKOuvert then Exit ;
s:='' ;
Q:=OpenSQL('SELECT J_CONTREPARTIE from JOURNAL Where J_JOURNAL="'+ Journal +'"',TRUE) ;
if Not Q.Eof then
  BEGIN
  s:=Q.FindField('J_CONTREPARTIE').AsString;
  END ;
Ferme(Q) ;
Result:=s ;
VH^.CPAppelOLE:=True ;
end;

Function THalleyWindow.GetModesaisieJrl (Journal :string) : Variant ;
var s : string ;
    Q : TQuery ;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
if Not V_PGI.OKOuvert then Exit ;
s:='' ;
Q:=OpenSQL('SELECT J_MODESAISIE from JOURNAL Where J_JOURNAL="'+ Journal +'"',TRUE) ;
if Not Q.Eof then
  BEGIN
  s:=Q.FindField('J_MODESAISIE').AsString;
  END ;
Ferme(Q) ;
Result:=s ;
VH^.CPAppelOLE:=True ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 13/07/2005
Modifié le ... : 13/07/2005
Description .. : Verbes pour la révision
Suite ........ : Ces verbes n'ont pas forcément un lien avec la comptabilité 
Suite ........ : mais cela évite à la révision d'interroger plusieurs serveurs 
Suite ........ : OLE.
Mots clefs ... : 
*****************************************************************}
Procedure RevStocker(pStNomChamps : String; pStValeur : String; var RevVal,RevNCh : string);
const
   cSeparateur = ';';
   cRetourChariot = chr(13);
begin
      pStNomChamps := Trim(Uppercase(pStNomChamps));

      if RevNCh <> '' then
      begin
         RevNCh := RevNCh + cSeparateur;
         RevVal := RevVal + cSeparateur;
      end
      else
         if RevVal <> '' then
            RevVal := RevVal + cRetourChariot;

      RevNCh := RevNCh + pStNomChamps;
      RevVal := RevVal + pStValeur;
end;

Function THalleyWindow.GetPROVCPRTT (lDateArrete: variant): Variant;
var
   aDateArrete :string;
   aAn, aMois, ajour : string;
   lObQuery       : TQuery;
   lStSQL         : String;
   RevValeurs,RevNomsChamps : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;

   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
   RevValeurs := '';
   RevNomsChamps:='';
//   "05/31/2005"
   aDateArrete:=lDateArrete;
   aAn:=copy(aDateArrete,7,4);
   aJour:=copy(aDateArrete,1,2);
   aMois := copy(aDateArrete,4,2);
   aDateArrete:= aMois + '/' + aJour + '/' + aAn;
   try
      lStSQL := ' SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM,' +
                '  ETB_LIBELLE,ETB_DATECLOTURECPN, ETB_NBJOUTRAV, PROVCP.* FROM PROVCP' +
                '  LEFT JOIN SALARIES ON PSA_SALARIE=PDC_SALARIE' +
                '  LEFT JOIN ETABCOMPL ON PDC_ETABLISSEMENT=ETB_ETABLISSEMENT' +
                '  WHERE PDC_DATEARRET= ' +  '"' + aDateArrete + '"' +
                '  ORDER BY PDC_ETABLISSEMENT';

      lObQuery := OpenSQL(lStSQL, True);

      While not lObQuery.Eof do
      begin
            RevNomsChamps:='';
            RevStocker('Date arrêté',   lObQuery.FindField('PDC_DATEARRET').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Etablissement',  lObQuery.FindField('PSA_ETABLISSEMENT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Nom Etablissement',  lObQuery.FindField('ETB_LIBELLE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date Clôture CP année en cours',  lObQuery.FindField('ETB_DATECLOTURECPN').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Nb de jours travaillés(ouvré/able)',  lObQuery.FindField('ETB_NBJOUTRAV').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Salarié',      lObQuery.FindField('PSA_SALARIE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Nom',      lObQuery.FindField('PSA_LIBELLE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Prénom',      lObQuery.FindField('PSA_PRENOM').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP base reliquat N-1',   lObQuery.FindField('PDC_CPBASERELN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP Mois reliquat N-1', lObQuery.FindField('PDC_CPMOISRELN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP restant reliquat N-1',       lObQuery.FindField('PDC_CPRESTRELN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP base N-1',   lObQuery.FindField('PDC_CPBASEN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP Mois N-1', lObQuery.FindField('PDC_CPMOISN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP acquis N-1',     lObQuery.FindField('PDC_CPACQUISN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP pris N-1',     lObQuery.FindField('PDC_CPPRISN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP restant N-1',     lObQuery.FindField('PDC_CPRESTN1').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP base N',   lObQuery.FindField('PDC_CPBASEN').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP Mois N', lObQuery.FindField('PDC_CPMOISN').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP acquis N',     lObQuery.FindField('PDC_CPACQUISN').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP pris N',     lObQuery.FindField('PDC_CPPRISN').AsString,RevValeurs,RevNomsChamps);
            RevStocker('CP restant N',     lObQuery.FindField('PDC_CPRESTN').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Provision CP',     lObQuery.FindField('PDC_CPPROVISION').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Coefficient de charges',     lObQuery.FindField('PDC_COEFFCHARGES').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Charges patronales CP',     lObQuery.FindField('PDC_CPCHARGESPAT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Nombre de RTT restants',     lObQuery.FindField('PDC_NBRTT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Taux journalier',     lObQuery.FindField('PDC_TAUXJOUR').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Provision RTT',     lObQuery.FindField('PDC_PROVRTT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Charges patronales RTT',     lObQuery.FindField('PDC_RTTCHARGESPAT').AsString,RevValeurs,RevNomsChamps);
            lObQuery.Next;
      end;

      // Résultat
      Result := RevNomsChamps + cRetourChariot + RevValeurs;
   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
   end;
end;

Function THalleyWindow.GetPREPDADS (lDateDeb, lDateFin: variant): Variant;
var
   aDatedeb,aDateFin :string;
   aAn, aMois, ajour : string;
   lObQuery       : TQuery;
   lStSQL         : String;
   RevValeurs,RevNomsChamps : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;

   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
   RevValeurs := '';
   RevNomsChamps:='';
//   "05/31/2005"
   aDateDeb:=lDateDeb;
   aDateFin:=lDateFin;
   aAn:=copy(lDateFin,7,4);
   aJour:=copy(lDateFin,1,2);
   aMois := copy(lDateFin,4,2);
   aDateFin:= aMois + '/' + aJour + '/' + aAn;

   aAn:=copy(lDateDeb,7,4);
   aJour:=copy(lDateDeb,1,2);
   aMois := copy(lDateDeb,4,2);
   aDateDeb:= aMois + '/' + aJour + '/' + aAn;

   try
      lStSQL := 'SELECT PPU_DATEDEBUT,PPU_DATEFIN,SUM(PPU_CHEURESTRAV) AS CHEURESTRAV, ' +
      'SUM(PPU_CBRUT) AS CBRUT, SUM(PPU_CBRUTFISCAL) AS CBRUTFISCAL,SUM(PPU_CPLAFONDSS) AS CPLAFONDSS, ' +
      'SUM(PPU_CCOUTPATRON) AS CCOUTPATRON,SUM(PPU_CCOUTSALARIE) AS CCOUTSALARIE, ' +
      'SUM(PPU_CNETIMPOSAB) AS CNETIMPOSAB,SUM(PPU_CNETAPAYER) AS CNETAPAYER ' +
      'FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + aDateDeb + '" AND PPU_DATEFIN<="' + aDateFin + '" ' +
      'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN ' +
      'ORDER BY PPU_DATEDEBUT,PPU_DATEFIN';

      lObQuery := OpenSQL(lStSQL, True);

      While not lObQuery.Eof do
      begin
            RevNomsChamps:='';
            RevStocker('Date début',   lObQuery.FindField('PPU_DATEDEBUT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date fin',  lObQuery.FindField('PPU_DATEFIN').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Brut',  lObQuery.FindField('CBRUT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Brut fiscal',  lObQuery.FindField('CBRUTFISCAL').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Plafond Sécu',  lObQuery.FindField('CPLAFONDSS').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Coût patronal',      lObQuery.FindField('CCOUTPATRON').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Coût salarial',      lObQuery.FindField('CCOUTSALARIE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Net imposable',      lObQuery.FindField('CNETIMPOSAB').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Net à payer',   lObQuery.FindField('CNETAPAYER').AsString,RevValeurs,RevNomsChamps);
            lObQuery.Next;
      end;

      // Résultat
      Result := RevNomsChamps + cRetourChariot + RevValeurs;
   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
   end;
end;

function THalleyWindow.GetCreditBails (LDateDeb: variant): Variant;
var
   aDateDeb :string;
   aAn, aMois, ajour : string;
   lObQuery       : TQuery;
   lStSQL         : String;
   RevValeurs,RevNomsChamps : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
   RevValeurs := '';
   RevNomsChamps:='';
//   "05/31/2005"
   aAn:=copy(lDateDeb,7,4);
   aJour:=copy(lDateDeb,1,2);
   aMois := copy(lDateDeb,4,2);
   aDateDeb:= aMois + '/' + aJour + '/' + aAn;

   try
      lStSQL := 'SELECT I_NUMCONTRATCB, I_LIBELLE, I_VERSEMENTCB, I_PERIODICITE, I_DATEDEBUTECH, '+
      'I_DATEFINECH, I_MONTANTSUIVECHE ' +
      'FROM IMMO ' +
      'WHERE I_NATUREIMMO="CB" AND  I_DATEFINECH >="' + aDateDeb + '" ' +
      'ORDER BY I_LIBELLE';

      lObQuery := OpenSQL(lStSQL, True);

      While not lObQuery.Eof do
      begin
            RevNomsChamps:='';
            RevStocker('Numéro contrat',   lObQuery.FindField('I_NUMCONTRATCB').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Libelle',  lObQuery.FindField('I_LIBELLE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Paiement',  lObQuery.FindField('I_VERSEMENTCB').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Périodicité',  lObQuery.FindField('I_PERIODICITE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date début',  lObQuery.FindField('I_DATEDEBUTECH').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date fin',      lObQuery.FindField('I_DATEFINECH').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Montant échéance',      lObQuery.FindField('I_MONTANTSUIVECHE').AsString,RevValeurs,RevNomsChamps);
            lObQuery.Next;
      end;

      // Résultat
      Result := RevNomsChamps + cRetourChariot + RevValeurs;
   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
   end;
End;

function THalleyWindow.GetJustImmos(LDateDeb: variant): Variant;
var   lObQuery       : TQuery;
   lStSQL         : String;
   RevValeurs,RevNomsChamps : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
  try
    lStSQL := 'SELECT I_IMMO,I_LIBELLE,I_COMPTEIMMO, I_DATEPIECEA, ' +
      'I_MONTANTHT,I_METHODEECO,I_TAUXECO, I_NATUREIMMO,I_ETAT '+
      'FROM IMMO WHERE I_NATUREIMMO="PRO" AND I_ETAT<>"FER" AND I_QUALIFIMMO="R" ' +
      'AND I_DATEPIECEA>="'+ USDatetime(LDateDeb) +'" '+ 'AND I_DATEPIECEA<="'+USDatetime(VH^.Encours.Fin)+'"';

      lObQuery := OpenSQL(lStSQL, True);

      While not lObQuery.Eof do
      begin
            RevNomsChamps:='';
            RevStocker('Numéro immo',   lObQuery.FindField('I_IMMO').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Libelle',  lObQuery.FindField('I_LIBELLE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Compte',  lObQuery.FindField('I_COMPTEIMMO').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date acquisition',  lObQuery.FindField('I_DATEPIECEA').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Montant HT',  lObQuery.FindField('I_MONTANTHT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Méthode Amt',      lObQuery.FindField('I_METHODEECO').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Taux Amt',      lObQuery.FindField('I_TAUXECO').AsString,RevValeurs,RevNomsChamps);
            lObQuery.Next;
      end;

      // Résultat
      Result := RevNomsChamps + cRetourChariot + RevValeurs;
   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
   end;
end;


function THalleyWindow.GetJustSorties(LDateDeb: variant): Variant;
var lCumAmt : double;
   lObQuery       : TQuery;
   lStSQL         : String;
   RevValeurs,RevNomsChamps : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
  try
    lStSQL := 'SELECT I_IMMO,I_LIBELLE,I_COMPTEIMMO,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT,'+
      'I_METHODEECO,IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES,IL_VOCEDEE'+
      ' FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE I_NATUREIMMO="PRO" AND I_QUALIFIMMO="R" AND IL_TYPEOP LIKE "CE%" '+
      ' AND IL_DATEOP>="'+USDateTime(LDateDeb)+'"'+
      ' AND IL_DATEOP<="'+USDatetime(VH^.Encours.Fin)+'"';

      lObQuery := OpenSQL(lStSQL, True);

      While not lObQuery.Eof do
      begin
            RevNomsChamps:='';
            RevStocker('Numéro immo', lObQuery.FindField('I_IMMO').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Libelle',lObQuery.FindField('I_LIBELLE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Compte', lObQuery.FindField('I_COMPTEIMMO').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date acquisition',  lObQuery.FindField('I_DATEPIECEA').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Montant HT',  lObQuery.FindField('IL_VOCEDEE').AsString,RevValeurs,RevNomsChamps);
            lCumAmt:= lObQuery.FindField('IL_CUMANTCESECO').AsFloat+lObQuery.FindField('IL_DOTCESSECO').AsFloat;
            RevStocker('Cumul Amt.',      FloatToStr(lCumAmt),RevValeurs,RevNomsChamps);
            RevStocker('Date cession',      lObQuery.FindField('IL_DATEOP').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Prix de vente',      lObQuery.FindField('IL_MONTANTCES').AsString,RevValeurs,RevNomsChamps);
            lObQuery.Next;
      end;

      // Résultat
      Result := RevNomsChamps + cRetourChariot + RevValeurs;
   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
   end;
end;
function THalleyWindow.GetImmosFIMVT(LDateDeb: variant): Variant;
var
   lObQuery       : TQuery;
   lStSQL         : String;
   RevValeurs,RevNomsChamps : string;
begin
Result := JaiLeDroit;
if Result <> '' then exit;
   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
  try
    lStSQL := 'SELECT I_IMMO,I_LIBELLE,I_COMPTEIMMO,I_OPECESSION,I_DATEPIECEA,I_NATUREIMMO,I_VALEURACHAT,'+
      'IL_DATEOP,IL_VOCEDEE,I_MONTANTHT'+
      ' FROM IMMOLOG LEFT JOIN IMMO  ON IL_IMMO=I_IMMO WHERE I_NATUREIMMO="FI" AND I_QUALIFIMMO="R"' +
      ' AND (IL_TYPEOP LIKE "CE%" OR IL_TYPEOP LIKE "AC%")'+
      ' AND IL_DATEOP>="'+USDateTime(LDateDeb)+'"'+
      ' AND IL_DATEOP<="'+USDatetime(VH^.Encours.Fin)+'"';

      lObQuery := OpenSQL(lStSQL, True);

      While not lObQuery.Eof do
      begin
            RevNomsChamps:='';
            RevStocker('Numéro immo',   lObQuery.FindField('I_IMMO').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Libelle',  lObQuery.FindField('I_LIBELLE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Compte',  lObQuery.FindField('I_COMPTEIMMO').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date acquisition',  lObQuery.FindField('I_DATEPIECEA').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Montant HT',  lObQuery.FindField('I_MONTANTHT').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Date cession',      lObQuery.FindField('IL_DATEOP').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Montant cédé',      lObQuery.FindField('IL_VOCEDEE').AsString,RevValeurs,RevNomsChamps);
            RevStocker('Valeur achat',      lObQuery.FindField('I_VALEURACHAT').AsString,RevValeurs,RevNomsChamps);
            lObQuery.Next;
      end;

      // Résultat
      Result := RevNomsChamps + cRetourChariot + RevValeurs;
   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
   end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2006
Modifié le ... :   /  /
Description .. : Verbe pour savoir si l'on gère la saisie de TVA sur le Dossier
Mots clefs ... :
*****************************************************************}
function THalleyWindow.EstTvaActivee: Variant;
begin
  Result := CPEstTvaActivee;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2006
Modifié le ... : 23/01/2006
Description .. :
Mots clefs ... :
*****************************************************************}
function THalleyWindow.GetMontantTVA( vStCodeTva : string ; vDateDebut, vDateFin : TDateTime) : Variant;
begin
  Result := CPGetMontantTVA(vStCodeTva, vDateDebut, vDateFin);
end;

////////////////////////////////////////////////////////////////////////////////

end.
