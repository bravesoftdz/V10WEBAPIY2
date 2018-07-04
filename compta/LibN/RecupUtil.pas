unit RecupUtil;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls,
  ExtCtrls, ComCtrls, HStatus, Ent1, HEnt1,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HCtrls, DB, HMsgBox,
{$IFNDEF EAGLSERVER}
  PGIExec,
{$ENDIF}
  UTOB;


Function FiniRecupInfoDossier : Boolean ;
Procedure CreeCptGenParamsoc(LeCpt,NatCpt,ChampParamSoc1,ChampParamSoc2 : String ;
                             Var ChampVH1 : String ; Ind : Integer) ;
Procedure CreeCptCollParamsoc ;
Procedure CreeCptAuxParamsoc(LeCpt,LeColl,NatCpt,ChampParamSoc1 : String ; Var ChampVH1 : String ; Ind : Integer) ;
procedure InitSocietePCL;
procedure InitEtablissement;
function VerifCoherenceExo ( DebExo, FinExo : TDateTime ) : boolean;
function EstAlpha (Chaine : string) : Boolean;
function EstBaseCommune : Boolean;
procedure MajParamSocDesImmos;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamSoc;


var
    HM : array [0..11] of string = (
          'Résultat',
          'Compte de bilan',
          'Collectif client divers',
          'Collectif fournisseur divers',
          'Collectif salarié divers',
          'Collectif divers',
          'Client Attente',
          'Fournisseur Attente',
          'Salarié attente',
          'Divers attente',
          'Solde bordereau (récup)',
          'Attente (récupération dossier)');

// Fiche 10479
procedure MajParamSocDesImmos;
type
  TParamSocCompte = record
      N : string; // SOC_NOM
      D : string; // SOC_DATA
      C : string; // Caractère de bourrage
    end;
  TParamSocImmo = array[1..28] of TParamSocCompte;
const
  PARAMSOC_IMMO : TParamSocImmo = (( N : 'SO_CPTEAMORTINF'; D : '28'; C : '0'),
    ( N : 'SO_CPTEAMORTSUP'; D : '283'; C : '9'),
    ( N : 'SO_CPTECBINF'; D : '612'; C : '0'),( N : 'SO_CPTECBSUP'; D : '612'; C : '9'),
    ( N : 'SO_CPTEDEPOTINF'; D : '275'; C : '0'),( N : 'SO_CPTEDEPOTSUP'; D : '275'; C : '9'),
    ( N : 'SO_CPTEDEROGINF'; D : '145'; C : '0'),( N : 'SO_CPTEDEROGSUP'; D : '145'; C : '9'),
    ( N : 'SO_CPTEDOTEXCINF'; D : '6871'; C : '0'),( N : 'SO_CPTEDOTEXCSUP'; D : '6871'; C : '9'),
    ( N : 'SO_CPTEDOTINF'; D : '6811'; C : '0'),( N : 'SO_CPTEDOTSUP'; D : '6811'; C : '9'),
    ( N : 'SO_CPTEEXPLOITINF'; D : '7811'; C : '0'),( N : 'SO_CPTEEXPLOITSUP'; D : '7811'; C : '9'),
    ( N : 'SO_CPTEFININF'; D : '26'; C : '0'),( N : 'SO_CPTEFINSUP'; D : '27'; C : '9'),
    ( N : 'SO_CPTEIMMOINF'; D : '2'; C : '0'),( N : 'SO_CPTEIMMOSUP'; D : '23'; C : '9'),
    ( N : 'SO_CPTELOCINF'; D : '613'; C : '0'),( N : 'SO_CPTELOCSUP'; D : '613'; C : '9'),
    ( N : 'SO_CPTEPROVDERINF'; D : '68725'; C : '0'),( N : 'SO_CPTEPROVDERSUP'; D : '68725'; C : '9'),
    ( N : 'SO_CPTEREPDERINF'; D : '78725'; C : '0'),( N : 'SO_CPTEREPDERSUP'; D : '78725'; C : '9'),
    ( N : 'SO_CPTEREPEXCINF'; D : '7871'; C : '0'),( N : 'SO_CPTEREPEXCSUP'; D : '7871'; C : '9'),
    ( N : 'SO_CPTEVACEDEEINF'; D : '675'; C : '0'),( N : 'SO_CPTEVACEDEESUP'; D : '675'; C : '9'));
var i : integer;
begin
  for i:=1 to High(PARAMSOC_IMMO) do
  begin
    if GetParamSoc(PARAMSOC_IMMO[i].N) = '' then
    SetParamSoc(PARAMSOC_IMMO[i].N,
      BourreLaDoncSurLesComptes(PARAMSOC_IMMO[i].D,PARAMSOC_IMMO[i].C));
  end;
end;


procedure CreeCptGenParamsoc(LeCpt, NatCpt, ChampParamSoc1,
  ChampParamSoc2: String; var ChampVH1: String; Ind: Integer);
Var Cpt : String ;
    Q : tQuery ;
    Oksoc : Boolean;
BEGIN
Oksoc := TRUE;
If (ChampParamSoc1<>'') and (GetParamSocSecur(ChampParamSoc1,'')<> '') Then
     Oksoc := ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL ="'+GetParamSocSecur(ChampParamSoc1,'')+'" ');
Cpt:=BourreEtLess(LeCpt,fbGene) ;
Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL ="'+Cpt+'" ',FALSE) ;
If Q.Eof and not(Oksoc) Then
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('G_GENERAL').AsString      :=Cpt ;
  Q.FindField('G_LIBELLE').AsString      :=TraduireMemoire(HM[Ind]) ;
  Q.FindField('G_ABREGE').AsString       :=Copy(TraduireMemoire(HM[Ind]),1,17) ;
  Q.FindField('G_NATUREGENE').AsString   :=NatCpt ;
  If (NatCpt='COC') Or (NatCpt='COF') Or (NatCpt='COS') Or (NatCpt='COD') Then Q.FindField('G_COLLECTIF').AsString:='X' ;
    Q.FindField('G_CREERPAR').AsString   :='IMP' ;
  Q.FindField('G_SENS').AsString         :='M' ;
  Q.Post ;
  END Else Cpt:=Q.FindField('G_GENERAL').AsString ;
Ferme(Q) ;
If (ChampParamSoc1<>'') and  (GetParamSocSecur(ChampParamSoc1,'') = '') then SetParamSoc(ChampParamSoc1,Trim(Cpt)) ;
If (ChampParamSoc2<>'') and  (GetParamSocSecur(ChampParamSoc2,'')= '') Then SetParamSoc(ChampParamSoc2,Trim(Cpt)) ;
If ChampVH1='' Then ChampVH1:=Trim(Cpt) ;
END ;

procedure CreeCptCollParamsoc;
Var
    CollCli,CollFou,CollSAl,CollDiv : String ;
    TC,TobL                         : TOB ;
    Q                               : tQuery ;
    St,St1                          : String ;
    i                               : Integer ;
BEGIN
CollCli:='' ; CollFou:='' ; CollSAl:='' ; CollDiv:='' ;
TC:=TOB.Create('',Nil,-1) ;
St:='SELECT G_GENERAL,G_NATUREGENE FROM GENERAUX WHERE G_COLLECTIF="X" ' ;
Q:=OpenSQL(St,TRUE) ;
TC.LoadDetailDB('CPT','','',Q,False,True) ;
Ferme(Q) ;
For i:=0 To  TC.Detail.Count-1 Do
  BEGIN
  TOBL:=TC.Detail[i] ;
  If TOBL.GetValue('G_NATUREGENE')='COC' Then BEGIN If CollCli='' Then CollCli:=TOBL.GetValue('G_GENERAL') ; END Else
  If TOBL.GetValue('G_NATUREGENE')='COF' Then BEGIN If CollFou='' Then CollFou:=TOBL.GetValue('G_GENERAL') ; END Else
  If TOBL.GetValue('G_NATUREGENE')='COS' Then BEGIN If CollSal='' Then CollSal:=TOBL.GetValue('G_GENERAL') ; END Else
  If TOBL.GetValue('G_NATUREGENE')='COD' Then BEGIN If CollDiv='' Then CollDiv:=TOBL.GetValue('G_GENERAL') END ;
  END ;
TC.Free ;
  St1 := GetParamSocSecur('SO_DEFCOLCLI','') ;
  If (CollCli = '') and (St1 = '') Then
  BEGIN
  CreeCptGenParamsoc('4119999999999','COC','SO_DEFCOLCLI','',St1,2) ;
  {$IFNDEF NOVH}
           VH^.DefautCli:=St1 ;
  {$ENDIF}
  END Else
  BEGIN
      if St1 = '' then
      begin
           SetParamSoc('SO_DEFCOLCLI',Trim(CollCli)) ;
          {$IFNDEF NOVH}
            VH^.DefautCli:=Trim(CollCli) ;
          {$ENDIF}
      end
      else
      CreeCptGenParamsoc('4119999999999','COC','SO_DEFCOLCLI','',St1,2) ;
  END ;

  St1:=GetParamSocSecur('SO_DEFCOLFOU','') ;
  If (CollFou = '') and (St1 = '')Then
  BEGIN
       CreeCptGenParamsoc('4019999999999','COF','SO_DEFCOLFOU','',St1,3) ;
    {$IFNDEF NOVH}
       VH^.DefautFou:=St1 ;
    {$ENDIF}
  END Else
  BEGIN
    if St1 = '' then
    begin
      SetParamSoc('SO_DEFCOLFOU',Trim(CollFou)) ;
    {$IFNDEF NOVH}
      VH^.DefautFou:=Trim(CollFou) ;
    {$ENDIF}
    end
    else
       CreeCptGenParamsoc(St1,'COF','SO_DEFCOLFOU','',St1,3) ;
  END ;
  St1 := GetParamSocSecur('SO_DEFCOLSAL','') ;
  If (CollSal='')  and (St1 = '') Then
  BEGIN
       CreeCptGenParamsoc('4219999999999','COS','SO_DEFCOLSAL','',St1,4) ;
      {$IFNDEF NOVH}
        VH^.DefautSal:=St1 ;
      {$ENDIF}
  END Else
  BEGIN
       if St1 = '' then
       begin
            SetParamSoc('SO_DEFCOLSAL',Trim(CollSal)) ;
            {$IFNDEF NOVH}
            VH^.DefautSal:=Trim(CollSal) ;
            {$ENDIF}
       end
       else
       CreeCptGenParamsoc(St1,'COS','SO_DEFCOLSAL','',St1,4) ;
  END ;
  St1:=GetParamSocSecur('SO_DEFCOLDIV','') ;
  If (CollDiv = '')  and (St1 = '') Then
  BEGIN
       CreeCptGenParamsoc('4679999999999','COD','SO_DEFCOLDIV','',St1,5) ;
       {$IFNDEF NOVH}
       VH^.DefautDivers:=St1 ;
       {$ENDIF}
  END Else
  BEGIN
       if St1 = '' then
       begin
            SetParamSoc('SO_DEFCOLDIV',Trim(CollDiv)) ;
            {$IFNDEF NOVH}
            VH^.DefautDivers:=Trim(CollDiv) ;
            {$ENDIF}
       end
       else
       CreeCptGenParamsoc(St1,'COD','SO_DEFCOLDIV','',St1,5) ;
  END ;

  St1:=GetParamSocSecur('SO_DEFCOLDDIV','') ;
  if St1 <> '' then
   CreeCptGenParamsoc(St1,'COD','SO_DEFCOLDDIV','',St1,5) ;
  St1:=GetParamSocSecur('SO_DEFCOLCDIV','') ;
  if St1 <> '' then
   CreeCptGenParamsoc(St1,'COD','SO_DEFCOLCDIV','',St1,5) ;

END ;

procedure CreeCptAuxParamsoc(LeCpt, LeColl, NatCpt,
  ChampParamSoc1: String; var ChampVH1: String; Ind: Integer);
Var Cpt : String ;
    Q : tQuery ;
    Oksoc : Boolean;
BEGIN
      Oksoc := TRUE;
      If (ChampParamSoc1<>'') and (GetParamSocSecur(ChampParamSoc1,'')<> '') Then
           Oksoc := ExisteSQL ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE ="'+GetParamSocSecur(ChampParamSoc1,'')+'" ');
      Cpt:=BourreEtLess(LeCpt,fbAux) ;
      Q:=OpenSQL('SELECT * FROM TIERS WHERE T_AUXILIAIRE ="'+Cpt+'" ',FALSE) ;
      If (Q.Eof) and not Oksoc Then
        BEGIN
        Q.Insert ; InitNew(Q) ;
        Q.FindField('T_AUXILIAIRE').AsString             :=Cpt ;
        Q.FindField('T_LIBELLE').AsString                :=TraduireMemoire(HM[Ind]) ;
        Q.FindField('T_ABREGE').AsString                 :=Copy(TraduireMemoire(HM[Ind]),1,8) ;
        Q.FindField('T_NATUREAUXI').AsString             :=NatCpt ;
        Q.FindField('T_CREERPAR').AsString               :='IMP' ;
        Q.FindField('T_COLLECTIF').AsString              :=LeColl ;
        Q.FindField('T_MODEREGLE').AsString              :='002' ;
        Q.FindField('T_REGIMETVA').AsString              :='FRA' ;
        Q.FindField('T_TVAENCAISSEMENT').AsString        :='TD' ;
        Q.FindField('T_TIERS').AsString                  :=Cpt ;
        Q.FindField('T_LETTRABLE').AsString              :='-' ;
        Q.Post ;
        END Else Cpt:=Q.FindField('T_AUXILIAIRE').AsString ;
      Ferme(Q) ;
      If (ChampParamSoc1<>'') and (GetParamSocSecur(ChampParamSoc1,'')= '') Then SetParamSoc(ChampParamSoc1,Trim(Cpt)) ;
      If ChampVH1='' Then ChampVH1:=Trim(Cpt) ;
END ;


Function FiniRecupInfoDossier : Boolean ;
Var Cpt     : String ;
    Q       : TQuery ;
    St1,St  : String ;
procedure DefautInit;
var
St : String ;
begin
        St:=GetParamSocSecur('SO_CLIATTEND','');
        CreeCptAuxParamsoc('CATTEN0000000',GetParamSocSecur('SO_DEFCOLCLI',''),'CLI','SO_CLIATTEND',St,6) ;
{$IFNDEF NOVH} VH^.TiersDefCli:=St ; {$ENDIF}
        St1:=GetParamSocSecur('SO_FOUATTEND','') ;
        if St1 = '' then
        CreeCptAuxParamsoc('FATTEN0000000',GetParamSocSecur('SO_DEFCOLFOU',''),'FOU','SO_FOUATTEND',St1,7)
        else
        CreeCptAuxParamsoc(St1,GetParamSocSecur('SO_DEFCOLFOU',''),'FOU','SO_FOUATTEND',St1,7);
{$IFNDEF NOVH} VH^.TiersDefFou:=St1 ; {$ENDIF}
        St1:=GetParamSocSecur('SO_SALATTEND','') ;
        if St1 = '' then
         CreeCptAuxParamsoc('FSATTEN0000000',GetParamSocSecur('SO_DEFCOLSAL',''),'SAL','SO_SALATTEND',St1,8)
         else
         CreeCptAuxParamsoc(St1,GetParamSocSecur('SO_DEFCOLSAL',''),'SAL','SO_SALATTEND',St1,8);
{$IFNDEF NOVH}  VH^.TiersDefSal:=St1 ; {$ENDIF}
        St1:=GetParamSocSecur('SO_DIVATTEND','');
        if St1 = '' then
        CreeCptAuxParamsoc('FDATTEN0000000',GetParamSocSecur('SO_DEFCOLDIV',''),'DIV','SO_DIVATTEND',St1,9)
        else
        CreeCptAuxParamsoc(St1,GetParamSocSecur('SO_DEFCOLDIV',''),'DIV','SO_DIVATTEND',St1,9);
{$IFNDEF NOVH}   VH^.TiersDefDiv:=St1 ; {$ENDIF}
end;
BEGIN
  Result:=TRUE ;
  St1:='' ;
  St := GetParamSocSecur('SO_OUVREBEN','');
  if St = '' then
  CreeCptGenParamsoc('1200000000000','DIV','SO_OUVREBEN','SO_FERMEBEN',St1,0)
  else
  CreeCptGenParamsoc(St,'DIV','SO_OUVREBEN','SO_FERMEBEN',St1,0);

  St := GetParamSocSecur('SO_OUVREPERTE','');
  if St = '' then
  CreeCptGenParamsoc('1290000000000','DIV','SO_OUVREPERTE','SO_FERMEPERTE',St1,0)
  else
  CreeCptGenParamsoc(St,'DIV','SO_OUVREPERTE','SO_FERMEPERTE',St1,0);

  St := GetParamSocSecur('SO_RESULTAT','');
  if St = '' then
  CreeCptGenParamsoc('1280000000000','DIV','SO_RESULTAT','',St1,0)
  else
  CreeCptGenParamsoc(St,'DIV','SO_RESULTAT','',St1,0);

  St1:=GetInfoCpta(fbGene).AxGenAttente ; CreeCptGenParamsoc('4718999999999','DIV','','',St1,10) ;
{$IFNDEF NOVH} VH^.Cpta[fbGene].AxGenAttente:=St1 ; {$ENDIF}
  St1:=GetParamSocSecur('SO_OUVREBIL','') ;
  if St1 = '' then
  CreeCptGenParamsoc('8900000000000','DIV','SO_OUVREBIL','SO_FERMEBIL',St1,1)
  else
  CreeCptGenParamsoc(St1,'DIV','SO_OUVREBIL','SO_FERMEBIL',St1,1);

{$IFNDEF NOVH} VH^.OuvreBil:=St1 ; {$ENDIF}
  St1:=GetInfoCpta(fbGene).Attente ; CreeCptGenParamsoc('4710000000000','DIV','SO_GENATTEND','',St1,7) ;
{$IFNDEF NOVH} VH^.Cpta[fbGene].Attente:=St1 ; {$ENDIF}
  CreeCptCollparamSoc ;
  if GetParamSocSecur ('SO_CPLIENGAMME','') = 'SI' then
  begin
       if ExisteSQl('SELECT * FROM CORRESP WHERE CR_TYPE="SIS" AND CR_LIBELLE like "0%"') then
       begin
        St1 := GetParamSocSecur('SO_CLIATTEND','');
        CreeCptAuxParamsoc('9ATTEN0000000',GetParamSocSecur('SO_DEFCOLCLI',''),'CLI','SO_CLIATTEND',St1,6) ;
{$IFNDEF NOVH} VH^.TiersDefCli:=St1 ; {$ENDIF}
        St1:=GetParamSocSecur('SO_FOUATTEND','') ; CreeCptAuxParamsoc('0ATTEN0000000',GetParamSocSecur('SO_DEFCOLFOU',''),'FOU','SO_FOUATTEND',St1,7) ;
{$IFNDEF NOVH} VH^.TiersDefFou:=St1 ; {$ENDIF}
        St1:=GetParamSocSecur('SO_SALATTEND','') ; CreeCptAuxParamsoc('0SATTEN0000000',GetParamSocSecur('SO_DEFCOLSAL',''),'SAL','SO_SALATTEND',St1,8) ;
{$IFNDEF NOVH}  VH^.TiersDefSal:=St1 ; {$ENDIF}
        St1:=GetParamSocSecur('SO_DIVATTEND','') ; CreeCptAuxParamsoc('0DATTEN0000000',GetParamSocSecur('SO_DEFCOLDIV',''),'DIV','SO_DIVATTEND',St1,9) ;
{$IFNDEF NOVH}   VH^.TiersDefDiv:=St1 ; {$ENDIF}
       end
       else
       if ExisteSQl('SELECT * FROM CORRESP WHERE CR_TYPE="SIS" AND CR_LIBELLE like "C%"') then
          DefautInit;
  end
  else
          DefautInit;
  If (Trim(GetInfoCpta(fbAxe1).Attente)='') Then
    Cpt:=BourreEtLess('9999999999999999',fbAxe1)
  Else Cpt:=Trim(GetInfoCpta(fbAxe1).Attente) ;
  Q:=OpenSQL('SELECT * FROM SECTION WHERE S_SECTION ="'+Cpt+'" ',FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNew(Q) ;
    Q.FindField('S_SECTION').AsString     :=Cpt ;
    Q.FindField('S_LIBELLE').AsString     :=TraduireMemoire(HM[11]) ;
    Q.FindField('S_ABREGE').AsString      :=Copy(TraduireMemoire(HM[11]),1,8) ;
    Q.FindField('S_AXE').AsString         :='A1' ;
    Q.FindField('S_CREERPAR').AsString    :='IMP' ;
    Q.FindField('S_SENS').AsString        :='M' ;
    Q.Post ;
  END Else Cpt:=Q.FindField('S_SECTION').AsString ;
  Ferme(Q) ;
  Q:=OPENSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="RTV" AND CC_CODE="FRA"',FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('CC_TYPE').AsString:='RTV' ;
    Q.FindField('CC_CODE').AsString:='FRA' ;
    Q.FindField('CC_LIBELLE').AsString:='France' ;
    Q.FindField('CC_ABREGE').AsString:='France' ;
    Q.FindField('CC_LIBRE').AsString:='' ;
    Q.Post ;
  END ;
  Ferme(Q) ;
  // fiche 10327
  Q:=OPENSQL('SELECT * FROM MODEPAIE '(*WHERE MP_MODEPAIE="DIV"*) ,FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('MP_MODEPAIE').AsString:='DIV' ;
    Q.FindField('MP_LIBELLE').AsString:='FRA' ;
    Q.FindField('MP_ABREGE').AsString:='France' ;
    Q.FindField('MP_ENCAISSEMENT').AsString:='MIX' ;
    Q.FindField('MP_CATEGORIE').AsString:='CHQ' ;
    Q.Post ;
  END ;
  Ferme(Q) ;

  // fiche 10327
  Q:=OPENSQL('SELECT * FROM MODEREGL'(* WHERE MR_MODEREGLE="002"*) ,FALSE) ;
  If Q.Eof Then
  BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('MR_MODEREGLE').AsString:='002' ;
    Q.FindField('MR_LIBELLE').AsString:='CHEQUE' ;
    Q.FindField('MR_ABREGE').AsString:='CHEQUE' ;
    Q.FindField('MR_APARTIRDE').AsString:='ECR' ;
    Q.FindField('MR_PLUSJOUR').AsInteger:=5 ;
    Q.FindField('MR_ARRONDIJOUR').AsString:='PAS' ;
    Q.FindField('MR_NOMBREECHEANCE').AsInteger:=1 ;
    Q.FindField('MR_SEPAREPAR').AsString:='QUI' ;
    Q.FindField('MR_MONTANTMIN').AsFloat:=999999 ;
    Q.FindField('MR_REMPLACEMIN').AsString:='002' ;
    Q.FindField('MR_MP1').AsString:='CHQ' ;
    Q.FindField('MR_TAUX1').AsFloat:=100 ;
    Q.FindField('MR_ESC1').AsString:='-' ;
    Q.Post ;
  END ;
  Ferme(Q) ;
{$IFNDEF NOVH}
  VH^.Cpta[fbAxe1].Attente:=Trim(Cpt) ;
{$ENDIF}
  if GetParamSocSecur('SO_CHADEB1','') = '' then
  begin
       Cpt:=BourreEtLess('60000000000000',fbGene) ; SetParamSoc('SO_CHADEB1',Cpt) ;
       {$IFNDEF NOVH} VH^.FCha[1].Deb:=Cpt ; {$ENDIF}
  end;
  if GetParamSocSecur('SO_CHAFIN1','') = '' then
  begin
       Cpt:=BourreEtLess('69999999999999',fbGene) ; SetParamSoc('SO_CHAFIN1',Cpt) ;
       {$IFNDEF NOVH} VH^.FCha[1].Fin:=Cpt ; {$ENDIF}
  end;
  if GetParamSocSecur('SO_PRODEB1','') = '' then
  begin
       Cpt:=BourreEtLess('70000000000000',fbGene) ; SetParamSoc('SO_PRODEB1',Cpt) ;
       {$IFNDEF NOVH} VH^.FPro[1].Deb:=Cpt ; {$ENDIF}
  end;
  if GetParamSocSecur('SO_PROFIN1','') = '' then
  begin                  // fiche 10523
       Cpt:=BourreEtLess('79999999999999',fbGene) ; SetParamSoc('SO_PROFIN1',Cpt) ;
       {$IFNDEF NOVH} VH^.FPro[1].Fin:=Cpt ; {$ENDIF}
  end;
  if GetParamSocSecur('SO_BILDEB1','') = '' then
  begin
       Cpt:=BourreEtLess('10000000000000',fbGene) ; SetParamSoc('SO_BILDEB1',Cpt) ;
       {$IFNDEF NOVH} VH^.FBil[1].Deb:=Cpt ; {$ENDIF}
  end;
  if GetParamSocSecur('SO_BILFIN1','') = '' then
  begin
       Cpt:=BourreEtLess('59999999999999',fbGene) ; SetParamSoc('SO_BILFIN1',Cpt) ;
       {$IFNDEF NOVH} VH^.FBil[1].Fin:=Cpt ; {$ENDIF}
  end;
  // Fiche 10479
  MajParamSocDesImmos;

END ;


procedure InitSocietePCL;
var
  OB, OB_DETAIL: TOB;
begin
  // Création de l'enregistrement
  OB := TOB.Create('La société', nil, -1);
  OB_DETAIL := TOB.Create('SOCIETE', OB, -1);
  OB_DETAIL.PutValue('SO_SOCIETE', GetParamSocSecur('SO_SOCIETE',''));
  OB_DETAIL.PutValue('SO_LIBELLE', GetParamSocSecur('SO_LIBELLE','')); // CA - 27/11/2001
  OB_DETAIL.PutValue('SO_ADRESSE1', GetParamSocSecur('SO_ADRESSE1',''));
  OB_DETAIL.PutValue('SO_ADRESSE2', GetParamSocSecur('SO_ADRESSE2',''));
  OB_DETAIL.PutValue('SO_ADRESSE3', GetParamSocSecur('SO_ADRESSE3',''));
  OB_DETAIL.PutValue('SO_CODEPOSTAL', GetParamSocSecur('SO_CODEPOSTAL',''));
  OB_DETAIL.PutValue('SO_VILLE', GetParamSocSecur('SO_VILLE',''));
  OB_DETAIL.PutValue('SO_PAYS', GetParamSocSecur('SO_PAYS',''));
  OB_DETAIL.PutValue('SO_TELEPHONE', GetParamSocSecur('SO_TELEPHONE',''));
  OB_DETAIL.PutValue('SO_FAX', GetParamSocSecur('SO_FAX',''));
  OB_DETAIL.PutValue('SO_TELEX', GetParamSocSecur('SO_TELEX',''));
  OB_DETAIL.PutValue('SO_SIRET', GetParamSocSecur('SO_SIRET',''));
  OB_DETAIL.PutValue('SO_APE', GetParamSocSecur('SO_APE',''));
  OB.InsertOrUpdateDB(True);
  OB.Free;
end;

procedure InitEtablissement;
var
  OB, OB_DETAIL, TEtab: TOB;
begin
  // Si il existe déjà un établissement dans la base, on ne crée pas d'établissement
  TEtab := TOB.Create('',nil,-1);
  TEtab.LoadDetailDB('ETABLISS','','ET_ETABLISSEMENT',nil,False);
  if TEtab.Detail.Count > 0 then
  begin
    SetParamSoc('SO_ETABLISDEFAUT', TEtab.Detail[0].GetValue('ET_ETABLISSEMENT'));
    TEtab.Free;
    exit;
  end;
  TEtab.Free;
  // Création de l'établissement par défaut
  OB := TOB.Create('Les établissements', nil, -1);
  OB_DETAIL := TOB.Create('ETABLISS', OB, -1);
  OB_DETAIL.PutValue('ET_ETABLISSEMENT', GetParamSocSecur('SO_SOCIETE',''));
  OB_DETAIL.PutValue('ET_SOCIETE', GetParamSocSecur('SO_SOCIETE',''));
  OB_DETAIL.PutValue('ET_LIBELLE', GetParamSocSecur('SO_LIBELLE',''));
  OB_DETAIL.PutValue('ET_ABREGE', Copy (GetParamSocSecur('SO_LIBELLE',''),1,17));
  OB_DETAIL.PutValue('ET_ADRESSE1', GetParamSocSecur('SO_ADRESSE1',''));
  OB_DETAIL.PutValue('ET_ADRESSE2', GetParamSocSecur('SO_ADRESSE2',''));
  OB_DETAIL.PutValue('ET_ADRESSE3', GetParamSocSecur('SO_ADRESSE3',''));
  OB_DETAIL.PutValue('ET_CODEPOSTAL', GetParamSocSecur('SO_CODEPOSTAL',''));
  OB_DETAIL.PutValue('ET_VILLE', GetParamSocSecur('SO_VILLE',''));
  OB_DETAIL.PutValue('ET_PAYS', GetParamSocSecur('SO_PAYS',''));
  OB_DETAIL.PutValue('ET_TELEPHONE', GetParamSocSecur('SO_TELEPHONE',''));
  OB_DETAIL.PutValue('ET_FAX', GetParamSocSecur('SO_FAX',''));
  OB_DETAIL.PutValue('ET_TELEX', GetParamSocSecur('SO_TELEX',''));
  OB_DETAIL.PutValue('ET_SIRET', GetParamSocSecur('SO_SIRET',''));
  OB_DETAIL.PutValue('ET_APE', GetParamSocSecur('SO_APE',''));
  OB_DETAIL.PutValue('ET_JURIDIQUE', GetParamSocSecur('SO_NATUREJURIDIQUE',''));
  OB.InsertOrUpdateDB(True);
  OB.Free;

  // Etablissement par défaut = société = établissement unique.
  SetParamSoc('SO_ETABLISDEFAUT', GetParamSocSecur('SO_SOCIETE',''));
end;

function VerifCoherenceExo ( DebExo, FinExo : TDateTime ) : boolean;
var TExo : TOB;
    i : integer;
    bok : boolean;
begin
  bOk := True;
  TExo := TOB.Create ('',nil,-1);
  try
    TExo.LoadDetailDB('EXERCICE','','EX_DATEDEBUT',nil,True);
    for i:=0 to TExo.Detail.Count - 1 do
    begin
      if ((DebExo >= TExo.Detail[i].GetValue('EX_DATEDEBUT')) and (DebExo <= TExo.Detail[i].GetValue('EX_DATEFIN'))) then
      begin
        if (DebExo = TExo.Detail[i].GetValue('EX_DATEDEBUT')) and (FinExo = TExo.Detail[i].GetValue('EX_DATEFIN')) then
        begin
          bOk := True;
        end else bOk := False;
        break;
      end else if ((DebExo < TExo.Detail[i].GetValue('EX_DATEDEBUT')) and (FinExo<> TExo.Detail[i].GetValue('EX_DATEDEBUT')-1)) then
        bOk := False
      else if ((DebExo < TExo.Detail[i].GetValue('EX_DATEDEBUT')) and (FinExo= TExo.Detail[i].GetValue('EX_DATEDEBUT')-1)) then
      begin
        bOk := True;
        break;
      end;
    end;
  finally
    TExo.Free;
  end;
  Result := bOk;
end;

function EstAlpha (Chaine : string) : Boolean;
var
j : integer;
begin
      EstAlpha := TRUE;
      for j:=1 To  length(Chaine) do
      begin
              if not (Chaine[j] in Alpha) then
              begin
                    if (Chaine[j] <> 'é') and (Chaine[j] <> 'ä') and  (Chaine[j] <> 'è') and
                    (Chaine[j] <> 'â') and (Chaine[j] <> 'ê') and  (Chaine[j] <> 'ç') and
                    (Chaine[j] <> 'à') and (Chaine[j] <> 'ù') and  (Chaine[j] <> 'û') and
                    (Chaine[j] <> 'ù') and (Chaine[j] <> 'ü') and  (Chaine[j] <> 'û') and
                    (Chaine[j] <> 'ë') and (Chaine[j] <> 'î') and  (Chaine[j] <> 'ï') and
                    (Chaine[j] <> 'ô') and (Chaine[j] <> 'ö') then
                    begin
                        EstAlpha := FALSE; break;
                    end;
              end;
      end;
end;
// car V_PGI.InBaseCommune ne marche pas
function EstBaseCommune : Boolean;
begin
if (v_pgi.DBName = v_pgi.DefaultSectionDBName) then Result := TRUE
else Result := FALSE
end;


end.
