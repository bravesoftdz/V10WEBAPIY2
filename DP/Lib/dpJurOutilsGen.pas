unit dpJurOutilsGen;

interface

uses
   {$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   {$ELSE}
   {$ENDIF}
   UTOB, Classes, HMsgBox, UYFileSTD, DpJurOutilsBlob, CBPPath;

const
  cstNonRenseignee : string = '##';

procedure TraiteUserCollJur(var laTob : TOB;NoOpe: string);
procedure TraiteUserCollDos(var laTob : TOB);

procedure InitChampsSupPer(var T :TOB ; Code : string;Qual : string = '');
function  GetNomPersonne(T : TOB) : string;
function  GetNomPersonneDB(sGuidPer_p : string) : string;
function  GetNomLongPersonne(T: TOB) : string;
procedure ChargeNomNomL(var TP : TOB;TS : TOB);
procedure ChargeInfoConjointPersonne(var TP : TOB);
function  GetDateNaisPersonne(T : TOB) : string;
procedure GereAccord(var TP : TOB);
procedure GereDepartement(var TP : TOB);
procedure GerePersonneAssociee(var TP : TOB);
procedure ChargePerAssociee(var TP : TOB;sGuidPer_p,RefPerAss,NomCode,ChampNom,ChampNomL,ChampQual : string);
function  ExtraitSufPersonne(T : TOB;Suf : string; var SufPer : string) : TOB;
procedure SetInfoPersonne(var T : TOB;Suf : string; var TypeChp : string;Valeur : Variant);
function  GetInfoPersonne(T : TOB; Suf : string;var TypeChp: string) : Variant;
function  RechDomPersonne(T : TOB;NomTablette,NomChamp : string) : string;
function  IsChampType(Pref,Suf : string; var Tipe : string) : boolean;
function  AffecteChampSelonType( sTypeChamp_p : string; vValeur_p : variant) : variant;
function  ChampBlob(Suf : string;var RacCh : string) : boolean;
function  GetValChampLiensOle(T : TOB ; NomChamp : string) : string;
procedure SetValChampLiensOle(var T : TOB ; NomChamp,ValChamp : string);
function  GetBlobFusion(Texte,TypeCh : string) : variant;

//function  GetNomDocToLoad(Q : TQuery) : string;


procedure RecupereRacineSuffixe(VarName : string; var Rac,Suf : string);

procedure CreationEvenementDoc(sCodeAct_p, sModule_p, sCodeOpe_p, sDossier_p, sNomDoc_p : string);
procedure UpdateEvenement(sCodeAct_p, sModule_p, sCodeOpe_p, sDossier_p : string);


function  ChampCommunANL_ANN(Suf : string) : boolean;
procedure GereInscription(var T : TOB);
function  AppelFonction_ExisteLien(T: TOB ; sGuidPer_p, rac : string) : boolean;
function  AppelFonction_Chaine(valeur : string) : string;
function  AppelFonction_PGCD(num,denom : string) : string;
Function  CalculPGCD(N1, N2 : integer) : integer;
function  AppelFonction_VoirChamp(sTable_p : string; sChamp_p : string; tsWhere_p, tsValeur_p : array of variant) : variant;
function  AppelFonction_GetBible(sModule_p, sCodeAct_p : string) : string;

function  GetBible(sModule_p, sCodeAct_p, sCodeOp_p, sCodeDos_p : string;
                  bModeDev_p : boolean;
                  var sMaquette_p, sDocResult_p, sCrit1_p, sCrit2_p : string) : integer;

function  VarDossier(sRacine_p : string) : boolean;
function  VarJuridique(sRacine_p : string) : boolean;
function  VarSociete(sRacine_p : string) : boolean;
function  VarOperation(sRacine_p : string) : boolean;
function  VarActe(sRacine_p : string) : boolean;
function  VarRubrique(sRacine_p : string) : boolean;
function  VarHistoOperation(sRacine_p : string) : boolean;
function  VarHistoActe(sRacine_p : string) : boolean;
function  VarHistoRubrique(sRacine_p : string) : boolean;
function  VarExercice(sRacine_p : string) : boolean;
function  VarAutreInfos(sRacine_p : string) : boolean;
function  VarLien(sRacine_p : string) : boolean;
function  VarEvenement(sRacine_p : string) : boolean;
function  VarBauxFonds(sRacine_p : string) : boolean;

function  VarAnnuBis(sRacine_p : string) : boolean;
function  VarDPFiscal(sRacine_p : string) : boolean;
function  VarDPSocial(sRacine_p : string) : boolean;
function  VarDPOrga(sRacine_p : string) : boolean;
function  VarTiers(sRacine_p : string) : boolean;
function  VarTiersCompl(sRacine_p : string) : boolean;
function  VarDPTabCompta(sRacine_p : string) : boolean;
function  VarDPTabPaie(sRacine_p : string) : boolean;
function  VarDPTabGenPaie(sRacine_p : string) : boolean;
function  VarDPControl(sRacine_p : string) : boolean;
function  VarProspects(sRacine_p : string) : boolean;
function  VarInterloc(sRacine_p : string) : boolean;
function  VarContact(sRacine_p : string) : boolean;

function  GetVarGuid(vValeur_p : variant; sSuffixe_p : string) : variant;
function  GetSufToGuid(sSuffixe_p : string) : string;
function  GetSufIsGuid(sSuffixe_p : string) : boolean;


//////////// IMPLEMENTATION ///////////
implementation

uses
   {$IFDEF VER150}
   Variants,
   {$ENDIF}
   windows, comctrls, extctrls, HEnt1, ParamSoc, FileCtrl, Math,
   sysutils, AGLUtilOle, NumConv, HCtrls, DpJurOutils, UJurOutils;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 05/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TraiteUserCollJur(var laTob : TOB; NoOpe : string);
var
  laCle : string;
begin
  with laTOB do
  begin
    AddChampSup('JUR_RESPSIGN',true);
    PutValue('JUR_RESPSIGN',RechDom('TTUTILISATEUR',GetValue('JUR_RESP'),false));
    AddChampSup('JUR_RESPFONCT',true);
    PutValue('JUR_RESPFONCT',GetValChamp('UTILISAT','US_FONCTION','US_UTILISATEUR="'+GetValue('JUR_RESP')+'"'));

    AddChampSup('JUR_COLLSIGN',true);
    PutValue('JUR_COLLSIGN',RechDom('TTUTILISATEUR',GetValue('JUR_COLL'),false));
    AddChampSup('JUR_COLLFONCT',true);
    PutValue('JUR_COLLFONCT',GetValChamp('UTILISAT','US_FONCTION','US_UTILISATEUR="'+GetValue('JUR_COLL')+'"'));

    AddChampSup('JUR_USRSIGN',true);
    PutValue('JUR_USRSIGN',RechDom('TTUTILISATEUR',GetValue('JUR_UTILISATEUR'),false));
    AddChampSup('JUR_USRFONCT',true);
    PutValue('JUR_USRFONCT',GetValChamp('UTILISAT','US_FONCTION','US_UTILISATEUR="'+GetValue('JUR_UTILISATEUR')+'"'));

    if NoOpe <> '' then
    begin
       AddChampSup('JUR_CODEOP',true);
       PutValue('JUR_CODEOP',NoOpe);

       AddChampSup('JUR_OPLIBELLE',true);
       laCle :='(JOP_CODEOP="'+NoOpe+'" and JOP_CODEDOS="'+laTOB.GetValue('JUR_CODEDOS')+'")';
       PutValue('JUR_OPLIBELLE',GetValChamp('JUDOSOPER','JOP_LIBELLE',laCle));
       //NCX 05/04/01
       AddChampSup('JUR_MODULE',true);
       PutValue('JUR_MODULE',GetValChamp('JUDOSOPER','JOP_MODULE',laCle));
       //
       AddChampSup('JUR_OPDESCRIPT',true);
       PutValue('JUR_OPDESCRIPT',GetBlobFusion(GetValChamp('JUDOSOPER','JOP_OPDESCRIPT',laCle),'A'));
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 05/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TraiteUserCollDos(var laTob : TOB);
begin
  with laTOB do
  begin
    AddChampSup('DOS_USRSIGN',true);
    PutValue('DOS_USRSIGN',RechDom('TTUTILISATEUR',GetValue('DOS_UTILISATEUR'),false));
    AddChampSup('DOS_USRFONCT',true);
    PutValue('DOS_USRFONCT',GetValChamp('UTILISAT','US_FONCTION','US_UTILISATEUR="'+GetValue('DOS_UTILISATEUR')+'"'));
  end;
end;

// création des champs supplémentaires liée aux information de l'ANNUAIRE d'une personne
procedure InitChampsSupPer(var T : TOB;Code,Qual : string);
var
   TOBBlobAnn : TOB;
begin
   T.AddChampSup('CODE',true);
   T.PutValue('CODE',Code);
   T.AddChampSup('QUAL',true);
   T.PutValue('QUAL',Qual);
   // calcul champs composés AddChampSup
   T.AddChampSup('NOM',true);
   T.PutValue('NOM',GetNomPersonne(T));
   T.AddChampSup('NOML',true);
   T.PutValue('NOML',GetNomLongPersonne(T));
   GereAccord(T);
   GereInscription(T);
   T.AddChampSup('NATLIB',true);
   T.PutValue('NATLIB',RechDomPersonne(T,'YYNATIONALITE','ANN_NATIONALITE'));
   GereDepartement(T);
   T.AddChampSup('ALPAYSLIB',true);
   T.PutValue('ALPAYSLIB',RechDomPersonne(T,'TTPAYS','ANN_PAYS'));
   T.AddChampSup('NAIS',true);
   T.Putvalue('NAIS',GetDateNaisPersonne(T));
   ChargeInfoConjointPersonne(T);
   T.AddChampSup('FORMELIB',true);
   T.Putvalue('FORMELIB',RechDomPersonne(T,'JUFORMELIB','ANN_FORME'));
   //NCX 22/02/01
   T.AddChampSup('FORMEBIB',true);
   T.Putvalue('FORMEBIB',RechDomPersonne(T,'JUFORMEBIB','ANN_FORME'));
   T.AddChampSup('CAPDEVLIB',true);
   T.Putvalue('CAPDEVLIB',RechDomPersonne(T,'JUDEVISE','ANN_CAPDEV'));
   T.AddChampSup('CAPDEVSY',true);
   T.Putvalue('CAPDEVSY',RechDomPersonne(T,'JUDEVISESYM','ANN_CAPDEV'));
   T.AddChampSup('EXEDEB',true);
   T.Putvalue('EXEDEB',RechDomPersonne(T,'JUMOISCLOTLIB','ANN_MOISCLOTURE'));
   T.AddChampSup('EXEFIN',true);
   T.Putvalue('EXEFIN',RechDomPersonne(T,'JUMOISCLOTABR','ANN_MOISCLOTURE'));
   GerePersonneAssociee(T);
// création de la TOB des "blob" de l'enregistrement Annuaire
// cette TOB est une fille de la tob TOBAnnuaire
// concerne la table LIENSOLE
   TOBBlobAnn := TOB.Create('LIENSOLE',T,-1);
// on charge tous les enreg correspondant à la clé : "ANN";;Code
// on récupérera ainsi les blobs de la fiche JURIDIQUE
   if Code<>'' then
      TOBBlobAnn.LoadDetailDB('LIENSOLE','"ANN";"' + Code + '";','',nil,false,false);
end;

function GetNomPersonne(T : TOB) : string;
var
   Nom : string;
begin
   Nom := '';
   if T.GetValue('ANN_PPPM')='PM' then
   begin
      if T.GetValue('ANN_NOM1')<>'' then Nom:=T.GetValue('ANN_NOM1');
      if T.GetValue('ANN_NOM2')<>'' then Nom:=Nom+' '+T.GetValue('ANN_NOM2');
      if T.GetValue('ANN_NOM3')<>'' then Nom:=Nom+' '+T.GetValue('ANN_NOM3');
   end
   else // if T.GetValue('ANN_PPPM')='PP' then => considérons que c'est la valeur par défaut
   begin
      if T.GetValue('ANN_CV')<>'' then Nom:=T.GetValue('ANN_CV');
      if T.GetValue('ANN_NOM2')<>'' then Nom:=Nom+' '+T.GetValue('ANN_NOM2');
      if T.GetValue('ANN_NOM1')<>'' then Nom:=Nom+' '+T.GetValue('ANN_NOM1');
   end;

   Result := Trim(Nom);
end;

function GetNomLongPersonne(T: TOB) : string;
var
  Q : TQuery;
  Select,Requete,Nee : string;
begin
  if (T.GetValue('ANN_PPPM')='PP') then
  begin
    Nee := '';
    if (T.GetValue('ANN_SEXE')='F') and (T.GetValue('ANN_NOM3')<>'') then
      Nee := ' née '+ T.GetValue('ANN_NOM3');
    result := GetNomPersonne(T)+nee;
  end
  else
  if T.GetValue('ANN_PPPM')='PM' then
  begin
    Select := 'select ANN_FORME,JFJ_FORMECOURTE ';
    Requete := Select+'FROM ANNUAIRE LEFT JOIN JUFORMEJUR ON JFJ_FORME=ANN_FORME '+
               'WHERE ANN_GUIDPER = "' + T.GetString('ANN_GUIDPER') + '"';
    Q := OpenSQL(Requete,true);
    if (Q=nil) or (Q.EOF) then begin Ferme(Q); exit; end;
    Q.First;
    result := Q.FindField('JFJ_FORMECOURTE').AsString+' '+GetNomPersonne(T);
    Ferme(Q);
  end;
end;

function RechDomPersonne(T : TOB;NomTablette,NomChamp : string) : string;
begin
  if T.GetValue(NomChamp)='' then result := ''
  else result := RechDom(NomTablette,T.GetValue(NomChamp),false);
end;

procedure ChargeNomNomL(var TP : TOB;TS : TOB);
begin
  TP.AddChampSup('NOM',true);
  TP.PutValue('NOM',GetNomPersonne(TS));
  TP.AddChampSup('NOML',true);
  TP.PutValue('NOML',GetNomLongPersonne(TS));
end;

procedure ChargeInfoConjointPersonne(var TP : TOB);
var
  T : TOB;
  sGuid_l : string;
begin
  TP.AddChampSup('CJCV',true);  TP.Putvalue('CJCV','');
  TP.AddChampSup('CJNOM1',true);  TP.Putvalue('CJNOM1','');
  TP.AddChampSup('CJNOM2',true);  TP.Putvalue('CJNOM2','');
  TP.AddChampSup('CJNOM3',true);  TP.Putvalue('CJNOM3','');
  TP.AddChampSup('CJNAIS',true);  TP.Putvalue('CJNAIS','');
  sGuid_l := TP.GetValue('ANN_GUIDCJ');
  if (sGuid_l = '0') or (sGuid_l = '') then exit;

  T := TOB.Create('ANNUAIRE',nil,-1);
  T.SelectDB('"' + sGuid_l + '"', nil, false);
  TP.Putvalue('CJCV',T.GetValue('ANN_CV'));
  TP.Putvalue('CJNOM1',T.GetValue('ANN_NOM1'));
  TP.Putvalue('CJNOM2',T.GetValue('ANN_NOM2'));
  TP.Putvalue('CJNOM3',T.GetValue('ANN_NOM3'));
  TP.Putvalue('CJNAIS',GetDateNaisPersonne(T));
  T.Free;
end;

function GetDateNaisPersonne(T : TOB) : string;
var
  accord,datenais,vnais,dnais : string;
begin
  result := '';   vnais := cstNonRenseignee; dnais  := cstNonRenseignee;
  if T.GetValue('ANN_PPPM')='PP' then
  begin
    if T.GetValue('ANN_SEXE')='F' then accord := 'e' else accord :='';
    if T.GetValue('ANN_VILLENAIS')<>'' then vnais := T.GetValue('ANN_VILLENAIS');
    if T.GetValue('ANN_DEPTNAIS')<>'' then dnais := T.GetValue('ANN_DEPTNAIS');
    if T.GetValue('ANN_DATENAIS')=iDate1900 then datenais := cstNonRenseignee
    else datenais := FormateDate('j mmmm AAAA',T.GetValue('ANN_DATENAIS'));

    result := 'né' + accord + ' le ' + datenais + ' à ' + vnais + ' (' + dnais + ')';
  end;
end;

procedure GereAccord(var TP : TOB);
var
  Q : TQuery;
  AccordE,ProMin,ProMaj,Requete : string;
begin
  AccordE := ''; ProMin := ''; ProMaj := '';
  TP.AddChampSup('E',true);
  TP.Putvalue('E',AccordE);
  TP.AddChampSup('PROMIN',true);
  TP.Putvalue('PROMIN',ProMin);
  TP.AddChampSup('PROMAJ',true);
  TP.Putvalue('PROMAJ',ProMaj);

  if TP.GetValue('ANN_PPPM')='PP' then
  begin
    if TP.GetValue('ANN_SEXE')='M' then
    begin AccordE := ''; ProMin := 'il'; ProMaj := 'Il'; end
    else begin AccordE := 'e'; ProMin := 'elle'; ProMaj := 'Elle'; end;
  end
  else
  if TP.GetValue('ANN_PPPM')='PM' then
  begin
      ProMin := 'il'; ProMaj := 'Il';
      Requete := 'SELECT JFJ_ACCORDF,JFJ_FORME FROM JUFORMEJUR '+
               'WHERE JFJ_FORME="'+TP.GetValue('ANN_FORME')+'"';
      Q := OpenSQL(Requete,true);
      if (Q=nil) or (Q.EOF) then begin Ferme(Q); exit; end else Q.First;
      if Q.FindField('JFJ_ACCORDF').AsString='X' then
      begin ProMin := 'elle'; ProMaj := 'Elle'; AccordE := 'e';  end;
      Ferme(Q);
  end;
  TP.Putvalue('E',AccordE);
  TP.Putvalue('PROMIN',ProMin);
  TP.Putvalue('PROMAJ',ProMaj);
end;

procedure GereDepartement(var TP : TOB);
var
  Dpt,Dptp,Code : string;
begin
  Dpt:=''; Dptp := '';
  if TP.GetValue('ANN_PAYS')='FRA' then
  begin
    Code := Copy(TP.GetValue('ANN_ALCP'),0,2);
    if Code='97' then Code := Copy(TP.GetValue('ANN_ALCP'),0,3);
    if Code<>'' then
    begin
      Dpt := RechDom('JUDEPART',Code,false);
      Dptp := RechDom('JUDEPABR',Code,false)+RechDom('JUDEPART',Code,false);
    end;
  end;
  TP.AddChampSup('ALDPT',true);
  TP.Putvalue('ALDPT',Dpt);
  TP.AddChampSup('ALDPTP',true);
  TP.Putvalue('ALDPTP',Dptp);
  Dpt:=''; Dptp := '';
  if TP.GetValue('ANN_PAYS')='FRA' then
  begin
    Code := Copy(TP.GetValue('ANN_RMDEP'),0,2);
    if Code='97' then Code := Copy(TP.GetValue('ANN_RMDEP'),0,3);
    if Code<>'' then
    begin
      Dpt := RechDom('JUDEPART',Code,false);
      Dptp := RechDom('JUDEPABR',Code,false)+RechDom('JUDEPART',Code,false);
    end;
  end;
  TP.AddChampSup('RMDEPLIB',true);
  TP.Putvalue('RMDEPLIB',Dpt);
  TP.AddChampSup('RMDEPP',true);
  TP.Putvalue('RMDEPP',Dptp);
end;

procedure GerePersonneAssociee(var TP : TOB);
var sGuid_l : string;
begin
  sGuid_l := TP.GetValue('ANN_PERASS1GUID');
  ChargePerAssociee(TP,sGuid_l,'ANN_PERASS1','REPPPCODE','REPPPNOM','REPPPNOML','REPPPQUAL');
  sGuid_l := TP.GetValue('ANN_PERASS2GUID');
  ChargePerAssociee(TP,sGuid_l,'ANN_PERASS2','REPPMCODE','REPPMNOM','REPPMNOML','REPPMQUAL');
end;

procedure ChargePerAssociee(var TP : TOB;sGuidPer_p,RefPerAss,NomCode,ChampNom,ChampNomL,ChampQual : string);
var
  T : TOB;
  sQual : string;
begin
  T := nil;
  if (sGuidPer_p<>'0') and (sGuidPer_p <> '') then
  begin
    T := TOB.Create('ANNUAIRE',TP,-1);
    T.SelectDB('"' + sGuidPer_p + '"',nil,false);
  end;
  TP.AddChampSup(NomCode,true);
  TP.PutValue(NomCode,sGuidPer_p);
  TP.AddChampSup(ChampNom,true);
  if T<>nil then TP.PutValue(ChampNom,GetNomPersonne(T)) else TP.PutValue(ChampNom,'');
  TP.AddChampSup(ChampNomL,true);
  if T<>nil then TP.PutValue(ChampNomL,GetNomLongPersonne(T)) else TP.PutValue(ChampNomL,'');
  TP.AddChampSup(ChampQual,true);
  if T<>nil then sQual:=TP.GetValue(RefPerAss+'QUAL') else sQual := '';
  TP.PutValue(ChampQual,sQual);
  if T<>nil then T.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 02/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CreationEvenementDoc(sCodeAct_p, sModule_p, sCodeOpe_p, sDossier_p, sNomDoc_p : string);
var
   bExiste_l : boolean;
   sGuidPer_l, sWhere_l, sGuidEvt_l : string;
   bEtatPurge_l : boolean;
   OBEvenement_l, OBAction_l : TOB;
   dtDateTime_l : TDateTime;
begin
   OBAction_l := TOB.Create('JUBIBACTION', nil, -1);
   OBAction_l.LoadDetailFromSQL('SELECT JBA_DOCPURGE, JBA_EVTOPTIONNEL FROM JUBIBACTION ' +
                                'WHERE JBA_MODULE = "' + sModule_p + '" and JBA_CODEACT = "' + sCodeAct_p + '"', false);
   if OBAction_l.Detail.Count > 0 then
   begin
      bEtatPurge_l := OBAction_l.Detail[0].GetString('JBA_DOCPURGE') = 'X';
   end;
   OBAction_l.Free;

   // On ne crée par les alertes optionnelles

   sGuidPer_l := GetValChamp('JURIDIQUE','JUR_GUIDPERDOS','JUR_CODEDOS = "' + sDossier_p + '"');
   bExiste_l := ExisteSQL('SELECT * FROM JUEVENEMENT ' +
                          'WHERE JEV_CODEEVT = "DOC" AND JEV_DOMAINEACT = "JUR" ' +
                          '  AND JEV_EVTLIBABREGE = "' + sCodeAct_p + '" ' +
                          '  AND JEV_CODEDOS = "' + sDossier_p + '" ' +
                          '  AND JEV_CODEOP       = "' + sCodeOpe_p + '" ' +
                          '  AND JEV_GUIDPER = "' + sGuidPer_l + '"');

   OBEvenement_l := TOB.Create('JUEVENEMENT', nil, -1);

   sGuidEvt_l := AglGetGuid();
   OBEvenement_l.PutValue('JEV_GUIDEVT', sGuidEvt_l);
   OBEvenement_l.PutValue('JEV_NOEVT', -2);   
   OBEvenement_l.PutValue('JEV_USER1', V_PGI.USER);
   OBEvenement_l.PutValue('JEV_CODEEVT', 'DOC');
   OBEvenement_l.PutValue('JEV_DOMAINEACT', 'JUR');
   OBEvenement_l.PutValue('JEV_EVTLIBABREGE', sCodeAct_p);
   OBEvenement_l.PutValue('JEV_EVTLIBELLE', GetValChamp('JUDOSOPACT', 'JOA_LIBACT', 'JOA_MODULE = "' + sModule_p + '" and JOA_CODEACT = "' + sCodeAct_p + '"'));
   OBEvenement_l.PutValue('JEV_FAMEVT', GetValChamp('JUTYPEEVT', 'JTE_FAMEVT', 'JTE_CODEEVT = "DOC"'));
   OBEvenement_l.PutValue('JEV_FAIT', 'X');
   OBEvenement_l.PutValue('JEV_URGENCE', 0);
   OBEvenement_l.PutValue('JEV_ALERTE', '-');
   OBEvenement_l.PutValue('JEV_GUIDPER', sGuidPer_l); 
   OBEvenement_l.PutValue('JEV_CODEDOS', sDossier_p);
   OBEvenement_l.PutValue('JEV_CODEOP', sCodeOpe_p);

   dtDateTime_l := Date + Time;
   OBEvenement_l.PutValue('JEV_DATE', dtDateTime_l);
   OBEvenement_l.PutValue('JEV_DATEFIN', dtDateTime_l);
   OBEvenement_l.PutValue('JEV_DOCNOM', sNomDoc_p);
   OBEvenement_l.PutValue('JEV_DOCAPPLI', 'DOC');
   OBEvenement_l.PutValue('JEV_ETATDOC', 'REA');

   sWhere_l := 'JBA_MODULE = "' + sModule_p + '" and JBA_CODEACT = "' + sCodeAct_p + '"';
   if bEtatPurge_l then
      OBEvenement_l.PutValue('JEV_ETATPURGE', '1')
   else
      OBEvenement_l.PutValue('JEV_ETATPURGE', '0');

   OBEvenement_l.InsertDB(nil);
   OBEvenement_l.Free;

   if bExiste_l then
      ExecuteSQL('DELETE FROM JUEVENEMENT ' +
                'WHERE JEV_CODEEVT = "DOC" AND JEV_DOMAINEACT = "JUR" ' +
                '  AND JEV_EVTLIBABREGE = "' + sCodeAct_p + '" ' +
                '  AND JEV_CODEDOS = "' + sDossier_p + '" ' +
                '  AND JEV_CODEOP       = "' + sCodeOpe_p + '" ' +
                '  AND JEV_GUIDPER = "' + sGuidPer_l + '"' +
                '  AND JEV_GUIDEVT <> "' + sGuidEvt_l + '"');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Meriaux
Créé le ...... : 26/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure UpdateEvenement(sCodeAct_p, sModule_p, sCodeOpe_p, sDossier_p : string);
var
   OBBibAction_l : TOB;
   bDocPurge_l : boolean;
   sGuidEvt_l : string;
begin
   bDocPurge_l := false;
   OBBibAction_l := TOB.Create('JUBIBACTION', nil, -1);
   OBBibAction_l.LoadDetailDBFromSQL('JUBIBACTION',
                            'SELECT JBA_DOCPURGE ' +
                            'FROM JUBIBACTION ' +
                            'WHERE JBA_MODULE = "' + sModule_p + '" ' +
                            '  and JBA_CODEACT = "' + sCodeAct_p + '"');

   if OBBibAction_l.Detail.Count > 0 then
      bDocPurge_l := OBBibAction_l.Detail[0].GetString('JBA_DOCPURGE') = 'X';
   OBBibAction_l.Free;

   sGuidEvt_l := '';
   OBBibAction_l := TOB.Create('JUDOSOPACT', nil, -1);
   OBBibAction_l.LoadDetailDBFromSQL('JUDOSOPACT',
                            'SELECT JOA_GUIDEVT ' +
                            'FROM JUDOSOPACT ' +
                            'WHERE JOA_MODULE = "' + sModule_p + '" ' +
                            '  and JOA_CODEACT = "' + sCodeAct_p + '" ' +
                            '  and JOA_CODEOP = "' + sCodeOpe_p + '"');

   if OBBibAction_l.Detail.Count > 0 then
      sGuidEvt_l := OBBibAction_l.Detail[0].GetString('JOA_GUIDEVT');
   OBBibAction_l.Free;


   if sGuidEvt_l = '' then
      Exit;

   BeginTrans;
   try
      if bDocPurge_l then
         ExecuteSQL('DELETE FROM JUEVENEMENT WHERE JEV_GUIDEVT = "' + sGuidEvt_l + '"')
      else
         ExecuteSQL('UPDATE JUEVENEMENT ' +
                    'SET JEV_FAIT = "X", JEV_ETATDOC = "REA" ' +
                    'WHERE JEV_GUIDEVT = "' + sGuidEvt_l + '"');
   except
      V_PGI.IoError:=oeUnknown;
   end;

   if V_PGI.IoError=oeUnknown then
      RollBack
   else
      CommitTrans;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure RecupereRacineSuffixe(VarName : string; var Rac,Suf : string);
var
  iPos : integer;
begin
   iPos := Pos('_', VarName) - 1;
   Rac := Copy(VarName, 0, iPos);
   Suf := Copy(VarName, iPos + 2, Length(VarName) - iPos);
   Suf := GetSufToGuid(Suf);
end;

function VarDossier(sRacine_p : string) : boolean;
begin
  result := (sRacine_p = 'MDP') or (sRacine_p = 'DOS');
end;

function VarJuridique(sRacine_p : string) : boolean;
begin
  result := (sRacine_p = 'DOS') or (sRacine_p = 'JUR');
end;

function VarSociete(sRacine_p : string) : boolean;
begin
  result := (sRacine_p = 'STE') or (sRacine_p = 'ANN');
end;

function VarOperation(sRacine_p : string) : boolean;
begin
  result := sRacine_p = 'DOP';
end;

function VarRubrique(sRacine_p : string) : boolean;
begin
  result := sRacine_p = 'DOR';
end;

function VarActe(sRacine_p : string) : boolean;
begin
  result := sRacine_p = 'DOA';
end;

function VarHistoOperation(sRacine_p : string) : boolean;
begin
  result := sRacine_p = 'HOP';
end;

function VarHistoRubrique(sRacine_p : string) : boolean;
begin
  result := sRacine_p='HOR';
end;

function VarHistoActe(sRacine_p : string) : boolean;
begin
  result := sRacine_p='HOA';
end;

function VarExercice(sRacine_p : string) : boolean;
begin
  result := sRacine_p='EXE';
end;

function VarLien(sRacine_p : string) : boolean;
begin
  result := sRacine_p='ANL';
end;

function VarEvenement(sRacine_p : string) : boolean;
begin
  result := sRacine_p = 'EVE';
end;

function VarBauxFonds(sRacine_p : string) : boolean;
begin
   result := sracine_p = 'JBF';
end;

function VarAutreInfos(sRacine_p : string) : boolean;
begin
  result := ExisteSQL('select JTI_CODEINFO from JUTYPEINFO where JTI_RACINE="' + sRacine_p + '"');
end;

function VarInterloc(sRacine_p : string) : boolean;
begin
   result := sracine_p = 'ANI';
end;

function VarContact(sRacine_p : string) : boolean;
begin
   result := sracine_p = 'CON';
end;

function VarAnnuBis(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'ANB';
end;

function VarDPFiscal(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'DFI';
end;

function VarDPSocial(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'DSO';
end;

function VarDPOrga(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'DOR';
end;

function VarTiers(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'T';
end;

function VarTiersCompl(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'YTC';
end;

function VarDPTabCompta(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'DTC';
end;

function VarDPTabPaie(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'DTP';
end;

function VarDPTabGenPaie(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'DT1';
end;

function VarDPControl(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'DCL';
end;

function VarProspects(sRacine_p : string) : boolean;
begin
   result := sRacine_p = 'RPR';
end;


function GetNomPersonneDB(sGuidPer_p : string) : string;
var
  T : TOB;
begin
  T := TOB.Create('ANNUAIRE',nil,-1);
  T.SelectDB('"' + sGuidPer_p + '"',nil,false);
  result := T.GetValue('ANN_NOMPER');
  T.Free;
end;

function ExtraitSufPersonne(T : TOB;Suf : string; var SufPer : string) : TOB;
var
  Pos : PChar;
  iPos,iFils : integer;
begin
// on cherche si le suffixe contient le mot clé LIE1
  Pos := StrPos(PChar(Suf),PChar('LIE1'));
  iPos := Pos-PChar(Suf);
  if iPos=0 then
  begin
// si oui
// on récupère le nom du champ à traiter : si NOMPERLIE1 on aura NOMPER
    SufPer := Copy(Suf,Length('LIE1')+1,Length(Suf)-Length('LIE1'));
// on récupère l'indice de la Tob contenant les infos de la personne LIE1
    iFils := T.Detail[0].GetValue('INDICELIE1');
// s'il y en a une, on accède à sa tob et on la retour à la fonction appelante
    if iFils<>-1 then result := T.Detail[iFils] else result:=nil;
    exit;
  end;
// on cherche si le suffixe contient le mot clé COOPT
  Pos := StrPos(PChar(Suf),PChar('COOPT'));
  iPos := Pos-PChar(Suf);
  if iPos=0 then
  begin
// si oui
// on récupère le nom du champ à traiter : si NOMPERCOOPT on aura NOMPER
    SufPer := Copy(Suf,Length('COOPT')+1,Length(Suf)-Length('COOPT'));
// on récupère l'indice de la Tob contenant les infos de la personne COOPT
    iFils := T.Detail[0].GetValue('INDICECOOPT');
// s'il y en a une, on accède à sa tob et on la retour à la fonction appelante
    if iFils<>-1 then result := T.Detail[iFils] else result:=nil;
    exit;
  end;
// sinon, il s'agit de la personne directement rattachée au lien
// suffixe est NOMPER, on renvoi NOMPER, sans changement
  SufPer := Suf;
  result := T.Detail[0];
end;

procedure SetInfoPersonne(var T : TOB;Suf : string; var TypeChp : string;Valeur : Variant);
var
  RacCh,SufPer : string;
//  TT : TOB;
begin
  T := ExtraitSufPersonne(T,Suf,SufPer);
  if T<>nil then
  begin
    if ChampBlob(SufPer,RacCh) then SetValChampLiensOle(T,'OLE_'+SufPer,Valeur)
    else if IsChampType('ANN',SufPer,TypeChp) then
    begin
      T.PutValue('ANN_'+SufPer,Valeur);
    end
    else T.PutValue(SufPer,Valeur);
  end;
end;

function GetInfoPersonne(T : TOB; Suf : string;var TypeChp: string) : Variant;
var
  RacCh,SufPer : string;
begin
  T := ExtraitSufPersonne(T,Suf,SufPer);
  if T<>nil then
  begin
    if ChampBlob(SufPer,RacCh) then result := GetValChampLiensOle(T,'OLE_'+SufPer)
    else if IsChampType('ANN', SufPer, TypeChp) then
    begin
//      result := GetVarGuid(T, 'ANN_', SufPer);
      result := T.GetValue('ANN_' + SufPer);
    end
    else result := T.GetValue(SufPer);
  end
  else result := '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

{procedure GetListeEve(JlpTob : TOB;Rac,Suf : string; var List : TStringList);
var
  I : integer;
begin
  for I:=0 to JlpTob.Detail.Count-1 do
    List.Add(JlpTob.Detail[I].GetValue('JEV_' + Suf));
end;
procedure GetListeExe(JlpTob : TOB;Rac,Suf : string; var List : TStringList);
var
  I : integer;
begin
  for I:=0 to JlpTob.Detail.Count-1 do
    List.Add(JlpTob.Detail[I].GetValue('JDE_'+Suf));
end;

procedure GetListeOpe (JlpTob : TOB;Rac,Suf : string; var List : TStringList);
var
  I : integer;
begin
  for I:=0 to JlpTob.Detail.Count-1 do
    List.Add(JlpTob.Detail[I].GetValue('JOP_'+Suf));
end;

procedure GetListeAct (JlpTob : TOB;Rac,Suf : string; var List : TStringList);
var
  I : integer;
begin
  for I:=0 to JlpTob.Detail.Count-1 do
    List.Add(JlpTob.Detail[I].GetValue('JOA_'+Suf));
end;

procedure GetListeRub (JlpTob : TOB;Rac,Suf : string; var List : TStringList);
var
  I : integer;
begin
  for I:=0 to JlpTob.Detail.Count-1 do
    List.Add(JlpTob.Detail[I].GetValue('JOR_'+Suf));
end;}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function VarLienPersonne(Rac : string ) : boolean;
begin
  result := ExisteSQL('select JTF_FONCTION from JUTYPEFONCT where JTF_RACINE="'+Rac+'"');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function IsChampType(Pref,Suf : string; var Tipe : string) : boolean;
begin
  result := false;
  if (Pref='ANL') and (ChampCommunANL_ANN(Suf)) then // CAS DES CHAMPS EN COMMUN ....!!!!
  begin                                              // ANNUAIRE et ANNULIEN
    Tipe := ChampToType('ANN_'+Suf);
    exit;
  end
  else
  if (Pref='JDE') and ( (Suf='DL') or (Suf='DS') ) then
  begin
    result:=true;
    exit;
  end;
  result := ChampToNum(Pref+'_'+Suf)<>-1;
  if result then Tipe := ChampToType(Pref+'_'+Suf) else result := Suf='';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 05/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function AffecteChampSelonType( sTypeChamp_p : string; vValeur_p : variant ) : variant;
var
   vValeur_l : variant;
begin
   if sTypeChamp_p = 'DATE' then
      vValeur_l := VarAsType(vValeur_p, varDate)
   else if sTypeChamp_p = 'BOOLEAN' then
      vValeur_l := TraduitBoolean(vValeur_p)
   else if (sTypeChamp_p = 'COMBO') or (Copy(sTypeChamp_p, 1, 7) = 'VARCHAR') then
      vValeur_l := VarAsType(vValeur_p, varString)
   else if sTypeChamp_p = 'INTEGER' then
      vValeur_l := VarAsType(vValeur_p, varInteger)
   else if sTypeChamp_p = 'DOUBLE' then
      vValeur_l := VarAsType(vValeur_p, varDouble)
   else if sTypeChamp_p = 'BLOB' then
      vValeur_l := GetBlobFusion(vValeur_p, 'A')
   else
      vValeur_l := vValeur_p;
   result := vValeur_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ChampCommunANL_ANN(Suf : string) : boolean;
begin
  result := (Suf='GUIDPER') or (Suf='NOMPER') or (Suf='FORME');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ChampBlob(Suf : string;var RacCh : string) : boolean;
begin
  RacCh:='';
  result := Suf='NOTEPER';
  if not result then result := Suf='REGMATTXT';
  if not result then result := Suf='BLOCNOTES';
  if not result then result := Suf='ACTIVITE';
  if not result then result := Suf='OBJETSOC';
  if result then begin RacCh := 'OLE'; exit; end;
  if not result then result := Suf='NOTELIEN';
  if result then begin RacCh := 'ANL'; exit; end;
  if not result then result := Suf='NOTEDOS';
  if result then begin RacCh := 'JDO'; exit; end;
  if not result then result := Suf='NOTEEVT';
  if result then begin RacCh := 'JEV'; exit; end;
  result := false;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetValChampLiensOle(T : TOB ; NomChamp : string) : string;
var
  TT : TOB;
  sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l : string;
begin
   if NomChamp = 'OLE_NOTEPER' then
      NomChamp := 'OLE_BLOCNOTES';

   BlobGetCode(NomChamp, sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);

   TT := T.FindFirst(['LO_EMPLOIBLOB'], [sEmploiBlob_l], true);
//  TT := T.FindFirst(['LO_IDENTIFIANT'],[NomChamp],true);
   if TT <> nil then
      result := GetBlobFusion(TT.GetValue('LO_OBJET'),'A')
   else
      result := '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure SetValChampLiensOle(var T : TOB ; NomChamp,ValChamp : string);
var TT : TOB;
begin
// !!! a enlever si dans .DOC on change NOTEPER par BLOCNOTES
  if NomChamp='OLE_NOTEPER' then NomChamp:='OLE_BLOCNOTES';
  TT := T.FindFirst(['LO_IDENTIFIANT'],[NomChamp],true);
  if TT<>nil then TT.PutValue('LO_OBJET',ValChamp);
end;

{***********A.G.L.***********************************************
Auteur  ...... : E. PLIEZ
Créé le ...... : 27/01/2000
Modifié le ... :   /  /
Description .. : Renvoi le texte d'un blob sans les codes rtf . Nécessite l'utilisation d'un control TRichEdit
Mots clefs ... : RECUPERER TEXTE RICHEDIT RTF
*****************************************************************}
function GetBlobFusion(Texte,TypeCh : string) : variant;
var
   Lignes : HTStrings;
   pos : integer;
   Retour,Chaine : string;
   RichEdit: TRichEdit;
   Panel: TPanel;
begin
   Panel := TPanel.Create( nil );
   Panel.Visible := False;
   Panel.ParentWindow := GetDesktopWindow;
   RichEdit := TRichEdit.Create(Panel);
   RichEdit.Parent := Panel;
   Lignes := HTStringList.Create;
   Lignes.Text := Texte;
   StringsToRich(RichEdit,Lignes);
   Chaine := RichEdit.Text;
   if (Length(Chaine)>0 ) and (Chaine[Length(Chaine)]=#10) and (Chaine[Length(Chaine)-1]=#13) then
      pos := 2
   else
      pos := 0;
   Retour := Copy(Chaine, 0, Length(Chaine) - pos);

   if TypeNombre(TypeCh) then
      result := Valeur(Retour)
   else if TypeDate(TypeCh) then
      result := TraduitDate(Retour)
   else
      result := Retour;

   Lignes.Free;
   RichEdit.Free;
   Panel.Free;
end;

procedure GereInscription(var T : TOB);
var val, sSiren_l : string;
begin
   sSiren_l := CodeSiren( T.GetValue('ANN_SIREN') );
   T.PutValue('ANN_SIREN', sSiren_l);

   T.AddChampSup('RCSNO',TRUE);
   Val:='';
   if ( sSiren_l <> '' ) and ( T.GetVAlue('ANN_RCSVILLE' ) <> '' ) then
      Val := sSiren_l + ' RCS '+ T.GetVAlue('ANN_RCSVILLE');
   T.PutValue('RCSNO',Val);

   T.AddChampSup('RMNO',TRUE);
   Val:='';
   if ( sSiren_l <> '' ) and ( T.GetVAlue('ANN_RMDEP') <> '' ) then
      Val := sSiren_l + ' RM '+ T.GetVAlue('ANN_RMDEP');
   T.PutValue('RMNO',Val);
end;

function AppelFonction_ExisteLien(T: TOB ; sGuidPer_p, rac : string) : boolean;
begin
  result := T.FindFirst(['ANL_GUIDPER','ANL_RACINE'],[sGuidPer_p,rac],true)<>nil
end;

function AppelFonction_Chaine(valeur : string) : string;
begin
  result := valeur;
end;


Function AppelFonction_PGCD(num, denom :string) : string;
var
D1,D2 : Double;
N1,N2 : LongInt;
retour : integer;
begin
  D1 := Valeur(num);
  D2 := Valeur(denom);
  If (Frac(D1)<>0) or (Frac(D2)<>0) then result := '-1'
  else
   begin
    N1 := Trunc(D1);
    N2 := Trunc(D2);
    If N1<>N2 then
     begin
      Retour := CalculPGCD(N1,N2); //attention on sait pas si appel de la fonction qui suit ou celle de HEnt1
      If Retour=1 then result := '-1' else result := IntToStr(Retour);
     end
    else result := '1';
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : CHARRAIX
Créé le ...... : 04/10/2001
Modifié le ... :   /  /    
Description .. : Permet de trouve le plus grand commun dénominateur.
Mots clefs ... : PGCD;DENOMINATEUR;FRACTION;QUOTIENT
*****************************************************************}
Function CalculPGCD(N1,N2 :integer) : integer;
//Rend le PGCD si calcul Ok
//Rend 0 si une ou les deux sont nulle(s)
//Rend 1 si aucun denominateur commun
Var
V1,V2,Temp : LongInt;
BEGIN
//Le PGCD ne tient pas compte des signes.
N1 := Abs(N1); N2 := Abs(N2);
//V1 doit contenir la valeur la plus grande.
V1 := Max(N1,N2); V2 := Min(N1,N2);
If (V1<>0) and (V2<>0) then
begin
  Try //Il se peut que les deux valeurs n'aient pas de dénominateur commun.
   begin
    WHILE NOT (V2=0) DO
      begin
        Temp:=V1;
        V1 := V2;
        V2 := temp MOD V2 ;//On élimine petit à petit les diviseurs communs.
      end;
   result := V1;
   end
  except //pas de dénominateur commun.
   result := 1;
  end;
end
else result := 0;
//Cette fonction a été crée dans L'AGL pour l'usage de tous donc on pourrait la supprimer
//et appeler directement la fonction qui ce trouve dans L'AGL
END;

function AppelFonction_VoirChamp( sTable_p : string; sChamp_p : string; tsWhere_p, tsValeur_p : array of variant ) : variant;
var
   nCpt_l : integer;
   sWhere_l : string;

begin
   result := false;
   for nCpt_l := 1 to tsWhere_p[0] do
   begin
      if nCpt_l > 1 then
         sWhere_l := sWhere_l + ' and ';
      sWhere_l := sWhere_l + tsWhere_p[nCpt_l] + ' = "' + tsValeur_p[nCpt_l] + '"';
   end;
   result := GetValChamp(sTable_p, sChamp_p, sWhere_l );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function AppelFonction_GetBible(sModule_p, sCodeAct_p : string ) : string;
var
   sMaquette_l, sDocResult_l, sCrit1_l, sCrit2_l : string;
begin
   if GetBible(sModule_p, sCodeAct_p, '', '',
               false, sMaquette_l, sDocResult_l, sCrit1_l, sCrit2_l) <> 0 then
      result := ''
   else
      result := sMaquette_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetBible(sModule_p, sCodeAct_p, sCodeOp_p, sCodeDos_p : string;
                  bModeDev_p : boolean;
                  var sMaquette_p, sDocResult_p, sCrit1_p, sCrit2_p : string) : integer;
var
   OBBible_l : TOB;
   sRequete_l, sPerso_l, sPredef_l, sFile_l : string;
   iErrCode_l : integer;
begin
   result := 0;
   if (sCodeDos_p <> '') then
   begin
      sRequete_l := 'select JBA_DOCMAQ, JBA_DOCNOMCRE, JOA_DOCFAIT, JOA_DOCFAITDATE, ' +
                    '       JOA_DOCNOM, JBA_DOCPERSO, JOA_CODEDOS, JOA_CODEOP, JBA_PREDEFMAJ, ' +
                    '       JBA_CRIT1, JBA_CRIT2 ' +
                    'from JUBIBACTION, JUDOSOPACT ' +
                    'where JBA_MODULE = "' + sModule_p + '" ' +
                    '  and JBA_CODEACT = "' + sCodeAct_P + '" ' +
                    '  and JOA_CODEDOS = "' + sCodeDos_p + '" ';
      if (sCodeOp_p <> '') then
         sRequete_l := sRequete_l + '  and JOA_CODEOP = "' + sCodeOp_p + '"';
   end
   else
   begin
        sRequete_l := 'select JBA_DOCPERSO, JBA_DOCMAQ, JBA_PREDEFMAJ, JBA_CRIT1, JBA_CRIT2 ' +
                    'from JUBIBACTION ' +
                    'where JBA_MODULE = "' + sModule_p + '" ' +
                    '  and JBA_CODEACT = "' + sCodeAct_p + '"';
   end;

   OBBible_l := TOB.Create('', nil, -1);
   OBBible_l.LoadDetailFromSQL(sRequete_l);

   if OBBible_l.Detail.Count = 0 then
   begin
      PGIInfo('Module "' + sModule_p + '" et acte "' + sCodeAct_P + '" inconnu.', 'Fonction GETBIBLE');
      result := -1;
      OBBible_l.Free;
      Exit;
   end;


   // On récupère les critères d'accès du document
   sMaquette_p := OBBible_l.Detail[0].GetString('JBA_DOCMAQ');
   sPerso_l := String(OBBible_l.Detail[0].GetValue('JBA_DOCPERSO'));
   sPredef_l := OBBible_l.Detail[0].GetValue('JBA_PREDEFMAJ');

   GetDocToLoad(sMaquette_p, false, sMaquette_p, sCrit1_p, sCrit2_p, sPredef_l);
   if sCrit1_p = '' then
      sCrit1_p := OBBible_l.Detail[0].GetValue('JBA_CRIT1');
   if sCrit2_p = '' then
      sCrit2_p := OBBible_l.Detail[0].GetValue('JBA_CRIT2');

   sFile_l := TCBPPath.GetCegidUserTempPath + 'PGI\STD\' + V_PGI.CodeProduit + '\' + sCrit1_p + '\' + sCrit2_p + '\' + V_PGI.LangueParDefaut + '\' + sPredef_l + '\' + sMaquette_p;
   DeleteFile(PChar(sFile_l));

   iErrCode_l := AGL_YFILESTD_EXTRACT(sMaquette_p, V_PGI.CodeProduit, sMaquette_p,
                        sCrit1_p, sCrit2_p, '', '', '',
                        false, V_PGI.LangueParDefaut, sPredef_l);

   if iErrCode_l > 0 then
   begin
      PGIInfo(AGL_YFILESTD_GET_ERR(iErrCode_l) + ' : ' + sMaquette_p);
      result := -2;
      OBBible_l.Free;
      exit;
   end;

   // on construit le nom complet du document final (path complet)
   if (sCodeDos_p <> '') then
      sDocResult_p := GetNomDocToSave(sCodeDos_p, OBBible_l.Detail[0].GetString('JBA_DOCNOMCRE'), bModeDev_p, true);
   OBBible_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function GetVarGuid(vValeur_p : variant; sSuffixe_p : string) : variant;
begin
   if GetSufIsGuid(sSuffixe_p) then
   begin
      if (vValeur_p = '') or (vValeur_p = Null) then
         vValeur_p := '0';
   end;
   result := vValeur_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function GetSufToGuid(sSuffixe_p : string) : string;
begin
   if UpperCase(sSuffixe_p) = 'CODEBENEF' then sSuffixe_p := 'GUIDBENEF'
   else if UpperCase(sSuffixe_p) = 'CODECJ' then sSuffixe_p := 'GUIDCJ'
   else if UpperCase(sSuffixe_p) = 'CODEI' then sSuffixe_p := 'GUIDI'
   else if UpperCase(sSuffixe_p) = 'CODEPER' then sSuffixe_p := 'GUIDPER'
   else if UpperCase(sSuffixe_p) = 'CODEPERDOS' then sSuffixe_p := 'GUIDPERDOS'
   else if UpperCase(sSuffixe_p) = 'CODEX' then sSuffixe_p := 'GUIDX'
   else if UpperCase(sSuffixe_p) = 'COOPTCODE' then sSuffixe_p := 'COOPTGUID'
   else if UpperCase(sSuffixe_p) = 'CRSPCODE' then sSuffixe_p := 'CRSPGUID'
   else if UpperCase(sSuffixe_p) = 'MDRPCODE' then sSuffixe_p := 'MDRPGUID'
   else if UpperCase(sSuffixe_p) = 'NOACT' then sSuffixe_p := 'GUIDACT'
   else if UpperCase(sSuffixe_p) = 'NODP' then sSuffixe_p := 'GUIDPER'
   else if UpperCase(sSuffixe_p) = 'NOEVT' then sSuffixe_p := 'GUIDEVT'
   else if UpperCase(sSuffixe_p) = 'NOOGADP' then sSuffixe_p := 'GUIDPEROGA'
   else if UpperCase(sSuffixe_p) = 'NOPER' then sSuffixe_p := 'GUIDPERINT'
   else if UpperCase(sSuffixe_p) = 'NOREFTETEGRDP' then sSuffixe_p := 'GUIDPERTETEGRD'
   else if UpperCase(sSuffixe_p) = 'PERASS1CODE' then sSuffixe_p := 'PERASS1GUID'
   else if UpperCase(sSuffixe_p) = 'PERASS2CODE' then sSuffixe_p := 'PERASS2GUID'
   else if UpperCase(sSuffixe_p) = 'PERASS3CODE' then sSuffixe_p := 'PERASS3GUID';
   result := sSuffixe_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetSufIsGuid(sSuffixe_p : string) : boolean;
begin
   result := (UpperCase(sSuffixe_p) = 'GUIDBENEF') or (UpperCase(sSuffixe_p) = 'GUIDCJ') or
   (UpperCase(sSuffixe_p) = 'GUIDI') or (UpperCase(sSuffixe_p) = 'GUIDPER') or
   (UpperCase(sSuffixe_p) = 'GUIDPERDOS') or (UpperCase(sSuffixe_p) = 'GUIDX') or
   (UpperCase(sSuffixe_p) = 'COOPTGUID') or (UpperCase(sSuffixe_p) = 'CRSPGUID') or
   (UpperCase(sSuffixe_p) = 'MDRPGUID') or (UpperCase(sSuffixe_p) = 'GUIDACT') or
   (UpperCase(sSuffixe_p) = 'GUIDPER') or (UpperCase(sSuffixe_p) = 'GUIDEVT') or
   (UpperCase(sSuffixe_p) = 'GUIDPEROGA') or (UpperCase(sSuffixe_p) = 'GUIDPERINT') or
   (UpperCase(sSuffixe_p) = 'GUIDPERTETEGRD') or (UpperCase(sSuffixe_p) = 'PERASS1GUID') or
   (UpperCase(sSuffixe_p) = 'PERASS2GUID') or (UpperCase(sSuffixe_p) = 'PERASS3GUID');
end;

end.


